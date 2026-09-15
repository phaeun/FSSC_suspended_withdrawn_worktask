// tab-urgent.js — URGENT(정지·철회 임박) 섹션

import {TODAY, ddayChip, diff, fmt, getChecks, requestDashRefresh} from './core.js';

export function toggleUrgentSort(){
  window._urgentSort=!window._urgentSort;
  requestDashRefresh();
}

export function toggleUrgentDdayOnly(){
  window._urgentDdayOnly=!window._urgentDdayOnly;
  requestDashRefresh();
}


export function renderSuspendRows(rows){
  if(!rows.length)return'<tr class="empty"><td colspan="5">해당 없음</td></tr>';
  const checks=getChecks();
  // D+5까지 표시 (업로드 완료 포함 모두 유지) / [D-day] 전용보기 켜지면 정지 170일 경과 건 제외 D+n(기한 경과)은 숨김
  const filtered=rows.filter(r=>{
    const d=diff(r.expDate);
    if(d===null||!(d>=-5))return false;
    if(window._urgentDdayOnly){
      const suspendElapsed=r.suspendDate?Math.round((TODAY-r.suspendDate)/86400000):null;
      const isUrgentSuspend=suspendElapsed!==null&&suspendElapsed>=170;
      if(!isUrgentSuspend&&d<0)return false;
    }
    return true;
  });
  if(!filtered.length)return'<tr class="empty"><td colspan="5">해당 없음</td></tr>';
  const sorted=[...filtered].sort((a,b)=>{
    const da=diff(a.expDate)??999, db=diff(b.expDate)??999;
    return window._urgentSort?(db-da):(da-db);
  });
  return sorted.map(r=>{
    const c=checks[r.jobNo]||{};
    const urgSuspend=c.urgSuspend||0;
    const urgWithdraw=c.urgWithdraw||0;
    const urgReport=c.urgReport||false;
    const urgWithdrawDone=c.urgWithdrawDone||false;
    const urgPost2=c.urgPost2||false;
    const urgNotFirstPost=c.urgNotFirstPost||false;
    const urgUpload=c.urgUpload||false;
    const urgCertAudit=c.urgCertAudit||false;
    const monthUpload=c.monthUpload||false;
    const internal=c.internal||false;
    const chk1=c.chk1||false,chk2=c.chk2||false;
    let statusText='';
    if(urgUpload||monthUpload)statusText='<span class="chip cb">업로드 완료</span>';
    else if(urgCertAudit)statusText='<span class="chip cb">기한내심사</span>';
    else if(urgWithdraw===2)statusText='<span class="chip cg">철회완료</span>';
    else if(urgNotFirstPost)statusText='<span class="chip" style="background:#2a1a40;color:#b06aff;font-weight:700">갱신 이후</span>';
    else if(urgPost2)statusText='<span class="chip" style="background:#2a1a40;color:#b06aff;font-weight:700">사후2차</span>';
    else if(urgReport)statusText='<span class="chip cg" style="font-weight:700">보고서 접수</span>';
    else if(urgWithdraw===1)statusText='<span class="chip cr">철회처리 필요</span>';
    else if(urgSuspend===2)statusText='<span class="chip cg">정지완료</span>';
    else if(urgSuspend===1)statusText='<span class="chip co">정지처리 필요</span>';
    else if(internal)statusText='<span class="chip" style="background:#2a1a40;color:#b06aff;font-weight:700">내부심사</span>';
    else if(chk2)statusText='<span class="chip cg">심사주기 안내 2차</span>';
    else if(chk1)statusText='<span class="chip cb">심사주기 안내 1차</span>';
    const isThisYear=r.expDate&&r.expDate.getFullYear()===TODAY.getFullYear();
    // 올해(만료 임박) 업체는 업로드해야만 흐림 처리, 그 외 해에 만료되는 업체는 기한 내 심사만 해도 흐림 처리
    // 사후2차 또는 첫 사후 아님 체크시에도 흐림 처리
    const rowDone=(urgUpload||monthUpload||urgSuspend===2||urgWithdraw===2||urgPost2||urgNotFirstPost||(urgCertAudit&&!isThisYear));
    let rowStyle='';
    if(rowDone)rowStyle='opacity:.5';
    else if(isThisYear)rowStyle='background:#5c1414;color:#ff8a8a;font-weight:600';
    const suspendElapsed=r.suspendDate?Math.round((TODAY-r.suspendDate)/86400000):null;
    const isUrgentSuspend=suspendElapsed!==null&&suspendElapsed>=170;
    const dateCell=isUrgentSuspend
      ?`<span class="chip cr">정지 ${suspendElapsed}일 경과</span> ${fmt(r.suspendDate)}`
      :`${ddayChip(r.expDate)} ${fmt(r.expDate)}`;
    return`<tr onclick="openModal(${r.idx},'urgent')" style="${rowStyle}"><td title="${r.partner}">${r.partner||'-'}</td><td title="${r.nameKo}">${r.nameKo||r.nameEn||'-'}</td><td>${r.jobNo}</td><td>${dateCell}</td><td>${statusText}</td></tr>`;
  }).join('');
}

// onclick="..." 문자열에서 참조하는 이 파일의 함수들을 전역에 노출
Object.assign(window, {toggleUrgentSort, toggleUrgentDdayOnly});
