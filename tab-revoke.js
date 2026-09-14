// tab-revoke.js — 철회 예정 업체 섹션

import {TODAY, fmt, getChecks, getEmails, requestDashRefresh} from './core.js';

export function renderRevokeRows(rows){
  if(!rows.length)return'<tr class="empty"><td colspan="5">해당 없음</td></tr>';
  const checks=getChecks();
  // 공문 발송, 업로드 완료(인증 복구), 철회 완료된 것은 맨 아래로
  const sorted=[...rows].sort((a,b)=>{
    const ca=checks[a.jobNo]||{},cb=checks[b.jobNo]||{};
    const da=(ca.chk4||ca.revokeUpload||ca.revokeDone)?1:0, db=(cb.chk4||cb.revokeUpload||cb.revokeDone)?1:0;
    if(da!==db)return da-db;
    // 정지날짜 오래된 순 유지
    if(!a.suspendDate&&!b.suspendDate)return 0;
    if(!a.suspendDate)return 1;if(!b.suspendDate)return -1;
    return a.suspendDate-b.suspendDate;
  });
  return sorted.map(r=>{
    const c=checks[r.jobNo]||{};
    const el=r.suspendDate?Math.round((TODAY-r.suspendDate)/86400000):null;
    const cls=el===null?'cgr':el>=165?'cr':el>=100?'co':'cb';
    // 엑셀 녹색 강조4 기준
    // 80% 더 밝게 (심사주기 안내 1차 완료): #e2efda (light)
    // 40% 더 밝게 (공문 발송까지 / 업로드 완료 / 철회 완료): #a9d18e → 다크모드에선 어둡게 조정
    let rowBg='';
    const isLight=document.body.classList.contains('light');
    if(c.chk4||c.revokeUpload||c.revokeDone){
      // 공문 발송 완료, 업로드 완료(인증 복구), 또는 철회 완료 - 짙은 초록 (40% 더 밝게)
      rowBg=isLight?'background:#a9d18e;color:#1a3a1a':'background:#1a3a1a';
    }else if(c.chk3){
      // 심사주기 안내 완료 - 연두색 (80% 더 밝게)
      rowBg=isLight?'background:#e2efda;color:#1a3a1a':'background:#1a2e1a';
    }
    return`<tr onclick="openModal(${r.idx},'revoke')" style="${rowBg}">
      <td title="${r.partner}">${r.partner||'-'}</td>
      <td title="${r.nameKo}">${r.nameKo||r.nameEn||'-'}</td>
      <td>${r.jobNo}</td>
      <td>${fmt(r.suspendDate)}</td>
      <td><span class="chip ${cls}">${el!==null?el+'일 경과':'날짜없음'}</span></td>
    </tr>`;
  }).join('');
}

export function isSamePartner(a, b){
  if(!a||!b)return false;
  if(a===b)return true;
  // 별칭 매핑 확인
  const emails=getEmails();
  const getAliasGroup=(name)=>{
    for(const[k,entry]of Object.entries(emails)){
      const all=[k,...(entry.aliases||[])];
      if(all.some(x=>x===name))return all;
    }
    return[name];
  };
  const groupA=getAliasGroup(a);
  const groupB=getAliasGroup(b);
  // 두 그룹에서 같은 이름 있으면 true
  if(groupA.some(x=>groupB.includes(x)))return true;
  // 2글자 이상 겹치면 true
  const shorter=a.length<b.length?a:b;
  for(let i=0;i<=shorter.length-2;i++){
    const sub=shorter.slice(i,i+2);
    if(a.includes(sub)&&b.includes(sub))return true;
  }
  return false;
}


export function toggleRevokeDropdown(){
  const list=document.getElementById('revoke-filter-list');
  if(list)list.style.display=list.style.display==='none'?'block':'none';
}

export function selectRevokeFilter(val,label){
  window._revokeFilter=val;
  const btn=document.getElementById('revoke-filter-label');
  if(btn)btn.textContent=(val?label:'컨설팅사')+' ▽';
  const list=document.getElementById('revoke-filter-list');
  if(list)list.style.display='none';
  requestDashRefresh();
}

// onclick="..." 문자열에서 참조하는 이 파일의 함수들을 전역에 노출
Object.assign(window, {selectRevokeFilter, toggleRevokeDropdown});
