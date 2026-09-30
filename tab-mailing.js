// tab-mailing.js — 주기 안내 메일링 (월별/총), 메인 대시보드 FSSC/업무시트와는 완전히 별개로 동작

import {TODAY, addLog, buildMailingData, fmt, getChecks, getMailingLog, loadMailingFssc, loadMailingJob, mailingFsscName, mailingJobName, requestDashRefresh, saveChecks, setMailingField, setMailingNotice} from './core.js';

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
    return isLight?'background:#bcc5cf':'background:#232d38';
  return'';
}

// O를 선택하면 O만, X를 선택하면 X(+사유)만 보이고, 아직 아무것도 안 골랐을 때만 둘 다 보여줌
function buildNoticeCell(jobNo,noticeKey,notice,disabled){
  if(disabled){
    return`<td style="color:var(--muted);font-size:11px">1차 X라서 해당없음</td>`;
  }
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
        ${buildNoticeCell(r.jobNo,'notice2',ml.notice2||{},(ml.notice1||{}).mode==='x')}
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
  // 철회예정으로 바뀌면 메인 대시보드의 '요청 업체' 목록에도 뜨도록 반영
  if(val==='철회예정'){
    const c=getChecks();if(!c[jobNo])c[jobNo]={};
    c[jobNo].revoke=true;
    c[jobNo].revokeReason='주기안내(총)에서 철회예정 선택';
    saveChecks(c);
    requestDashRefresh();
  }
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


// onclick="..." 문자열에서 참조하는 이 파일의 함수들을 전역에 노출
Object.assign(window, {editMailingNoticeDate, exportMailingMonthExcel, exportMailingTotalExcel, loadMailingFssc, loadMailingJob, openMailingDetail, openMailingMailDraft, saveMailingCheck, saveMailingField, saveMailingProc, selectMailingProc, setMailingEmailFilter, setMailingMonthFilter, setMailingNoticeMode, setMailingNoticeStatus, switchMailingView, toggleMailingIntent, toggleMailingMonthCollapse, toggleMailingProcDropdown});
