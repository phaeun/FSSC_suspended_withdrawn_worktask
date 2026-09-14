// core.js — 상태, 저장소(Firestore/localStorage), 공용 유틸, 파일 업로드 파싱, 데이터 병합
// 이 파일은 다른 커스텀 모듈을 import하지 않는 최하위 계층입니다.

export const TODAY=(()=>{const d=new Date();d.setHours(0,0,0,0);return d;})();

export let fsscRaw=null,jobRaw=null,merged=[];

export let showHidden=false;

export let curSubTab='fssc'; // 'fssc' | 'fsms'


// ── 저장소 (Firestore 기반) ──

export const LS = {
  get(k){ try{ return JSON.parse(localStorage.getItem(k)); }catch(e){ return null; } },
  set(k,v){ try{ localStorage.setItem(k,JSON.stringify(v)); }catch(e){} },
  del(k){ localStorage.removeItem(k); }
};


// Firestore 동기 래퍼 (로컬 캐시 + 백그라운드 동기화)

export const FSCache = {};

export function fsSet(key, value) {
  FSCache[key] = value;
  LS.set(key, value);
  if(window.FS) window.FS.save(key, value); // 백그라운드 저장, await 불필요
}

export function fsGet(key) {
  if(FSCache[key] !== undefined) return FSCache[key];
  return LS.get(key); // 로컬 폴백
}


export const getEmails  = () => fsGet('emails2') || {};

export const saveEmails = o  => fsSet('emails2', o);

export const getChecks  = () => fsGet('checks')  || {};

export const saveChecks = o  => fsSet('checks', o);

export const getIgcStore= () => fsGet('igcStore')|| {};

export const saveIgcStore=o  => fsSet('igcStore', o);

export const getOverrides=() => fsGet('overrides')|| {};

export const saveOverrides=o => fsSet('overrides', o);


export function findEmail(p){
  if(!p)return '';
  const emails=getEmails();const pl=p.toLowerCase();
  for(const[name,entry]of Object.entries(emails)){
    const keys=[name,...(entry.aliases||[])].map(k=>k.toLowerCase());
    if(keys.some(k=>pl.includes(k)||k.includes(pl)))return entry.email||'';
  }return '';
}


export function parseDate(v){
  if(!v)return null;
  if(v instanceof Date)return isNaN(v)?null:v;
  if(typeof v==='number'){const d=new Date(Math.round((v-25569)*86400000));d.setMinutes(d.getMinutes()+d.getTimezoneOffset());return isNaN(d)?null:d;}
  const s=String(v).trim().replace(/\./g,'-');
  const m=s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if(m){const d=new Date(+m[1],+m[2]-1,+m[3]);return isNaN(d)?null:d;}
  return null;
}

export function fmt(d){if(!d)return '-';return `${d.getFullYear()}-${p2(d.getMonth()+1)}-${p2(d.getDate())}`;}

export function p2(n){return String(n).padStart(2,'0');}

export function diff(d){
  if(!d)return null;
  // 년도 무관 - 올해 기준으로 월/일만 비교
  const thisYear=new Date(TODAY.getFullYear(),d.getMonth(),d.getDate());
  return Math.round((thisYear-TODAY)/86400000);
}

export function ddayChip(d){
  const n=diff(d);if(n===null)return '';
  if(n<0)return`<span class="chip cb">D+${Math.abs(n)}</span>`;
  if(n===0)return`<span class="chip cr">D-day</span>`;
  if(n<=10)return`<span class="chip cr">D-${n}</span>`;
  if(n<=30)return`<span class="chip co">D-${n}</span>`;
  return`<span class="chip cg">D-${n}</span>`;
}

export function statusChip(s){
  const m={정상:'cg',유지:'cg',진행중:'cb',정지:'co',만료:'cr',취소:'cgr','취소(Withdrawn)':'cgr',진행취소:'cgr','정지(Suspended)':'co'};
  return s?`<span class="chip ${m[s]||'cgr'}">${s}</span>`:'';
}

export function normalizeStatus(s){
  if(!s)return '';
  if(s.includes('취소'))return '취소';
  if(s.includes('정지'))return '정지';
  if(s==='정상'||s==='유지')return '정상';
  return s;
}

export function setTxt(id,val){const el=document.getElementById(id);if(el)el.textContent=val;}

export function isOverseasCountry(country){
  if(!country)return false;
  const norm=country.toUpperCase().replace(/[.,]/g,' ').replace(/\s+/g,' ').trim();
  const domestic=['KOREA','KR','REPUBLIC OF KOREA','KOREA REPUBLIC OF','SOUTH KOREA','대한민국','한국'];
  return!domestic.includes(norm);
}


// ── 파일 로드 ──

export function readXlsx(file,cb){
  const r=new FileReader();
  r.onload=e=>cb(XLSX.read(e.target.result,{type:'binary',cellDates:true,raw:false}));
  r.readAsBinaryString(file);
}

export function loadFSSC(ev){
  const f=ev.target.files[0];if(!f)return;
  readXlsx(f,wb=>{
    const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
    if(!rows.length||!('Client number' in rows[0])){alert('FSSC 파일 오류: "Client number" 컬럼 필요');return;}
    fsscRaw=rows; fsSet('fsscRaw',rows); fsSet('fsscName',f.name);
    document.getElementById('lbl-fssc').textContent=f.name;
    if(jobRaw)mergeAndRender();
  });ev.target.value='';
}

export function loadJob(ev){
  const f=ev.target.files[0];if(!f)return;
  readXlsx(f,wb=>{
    const wsName=wb.SheetNames.find(n=>n.toLowerCase().includes('job'))||wb.SheetNames[0];
    const ws=wb.Sheets[wsName];
    const arr=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false});
    const hdr=arr[8]||[];
    const rows=arr.slice(9).map(row=>{
      const obj={};hdr.forEach((h,i)=>{if(String(h).trim())obj[String(h).trim()]=row[i]??'';});return obj;
    }).filter(o=>String(o['New Job No.']||'').trim());
    if(!rows.length){alert('업무시트 데이터 없음');return;}
    jobRaw=rows; fsSet('jobRaw',rows); fsSet('jobName',f.name);
    document.getElementById('lbl-job').textContent=f.name;
    if(fsscRaw)mergeAndRender();
  });ev.target.value='';
}

export function loadIGC(ev){
  const f=ev.target.files[0];if(!f)return;
  const r=new FileReader();
  r.onload=e=>{
    const doc=(new DOMParser()).parseFromString(e.target.result,'text/html');
    const rows=Array.from(doc.querySelectorAll('tr'));
    if(!rows.length){alert('ISO 파일 파싱 실패');return;}
    const hdr=Array.from(rows[0].querySelectorAll('td,th')).map(td=>td.textContent.trim());
    const colIdx={};hdr.forEach((h,i)=>colIdx[h]=i);
    const get=(row,key)=>{const cells=Array.from(row.querySelectorAll('td,th'));return colIdx[key]!==undefined?(cells[colIdx[key]]?.textContent.trim()||''):'';};
    const store=getIgcStore();let updated=0,added=0;
    rows.slice(1).forEach(row=>{
      const jobNo=get(row,'Job No.');if(!jobNo)return;
      const issuedDate=get(row,'승인일')||get(row,'최초발행일');
      const entry={jobNo,status:normalizeStatus(get(row,'상태')),nameKo:get(row,'업체명'),nameEn:get(row,'업체명(영문)'),coid:get(row,'인증서번호'),partner:get(row,'영업자'),email:get(row,'E-mail'),phone:get(row,'대표전화번호'),address:get(row,'위치'),issuedDate,expDate:get(row,'만료일'),country:get(row,'국가')||'KOREA',src:'IGC'};
      if(store[jobNo])updated++; else added++;
      store[jobNo]=entry;
    });
    saveIgcStore(store); fsSet('igcName',f.name);
    document.getElementById('lbl-igc').textContent=f.name;
    alert(`ISO 업데이트: ${added}개 추가, ${updated}개 업데이트`);
    mergeAndRender();
  };r.readAsText(f,'utf-8');ev.target.value='';
}

// ── MERGE ──

export function mergeAndRender(){
  if(!fsscRaw&&!jobRaw&&Object.keys(getIgcStore()).length===0){alert('파일을 업로드해 주세요.');return;}
  const igcList=Object.values(getIgcStore());
  const fsscList=[];
  if(fsscRaw&&jobRaw){
    const jobMap={};
    jobRaw.forEach(r=>{const k=String(r['New Job No.']||'').trim();if(k)jobMap[k]=r;});
    fsscRaw.forEach((r,i)=>{
      const jobNo=String(r['Client number']||'').trim();
      const job=jobMap[jobNo]||{};
      const partner=String(job['Partner']||'').trim();
      const expDate=parseDate(r['Valid until']||r['Expiry Date']);
      const initialCert=parseDate(r['Initial certification']);
      const fsscStatus=String(r['Status']||r['Certificate Status ']||r['Certificate Status']||'').trim();
      const compFull=String(job['Company name']||'').trim();
      const nameKo=compFull.split('/')[0].trim()||compFull;
      const issuedDate=parseDate(r['Issue date']||job['Issued Date']);
      fsscList.push({idx:0,jobNo,status:String(job['현황']||'').trim(),fsscStatus,partner,consultEmail:findEmail(partner),nameKo,nameEn:String(r['Organization']||'').trim(),coid:String(r['COID']||r['Certificate unique identification code']||'').trim(),certNo:String(job['Cert No.']||'').trim(),expDate,issuedDate,initialCert,suspendDate:parseDate(job['정지날짜']),email:String(r['Email of contact']||'').trim(),phone:String(r['Phonenumber of contact']||'').trim(),address:String(r['Address']||r['Street']||'').trim(),country:String(r['Country']||'KOREA').trim(),src:'FSSC'});
    });
  }
  const igcMapped=igcList.map(r=>({idx:0,jobNo:r.jobNo,status:r.status,fsscStatus:'',partner:r.partner,consultEmail:findEmail(r.partner),nameKo:r.nameKo,nameEn:r.nameEn,coid:r.coid,certNo:r.coid,expDate:parseDate(r.expDate),issuedDate:parseDate(r.issuedDate),suspendDate:null,email:r.email,phone:r.phone,address:r.address,country:r.country||'KOREA',src:'IGC'}));
  // 오버라이드 적용 (정지처리/철회처리)
  const overrides=getOverrides();
  merged=[...fsscList,...igcMapped].filter(r=>r.jobNo);
  merged.forEach((r,i)=>{
    r.idx=i;
    if(overrides[r.jobNo]){
      r._override=overrides[r.jobNo];
      if(overrides[r.jobNo]==='suspended')r.status='정지';
      if(overrides[r.jobNo]==='withdrawn')r.status='취소';
    }
  });
  requestFullRefresh();
}


// ── 미니 달력 ──

export let miniYear=TODAY.getFullYear(), miniMonth=TODAY.getMonth();


export function getMiniSrc(){ return curSubTab==='fssc'?'FSSC':'IGC'; }


export function renderMiniCalHtml(){
  const src=getMiniSrc();
  const mN=['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];
  const firstDay=new Date(miniYear,miniMonth,1).getDay();
  const lastDate=new Date(miniYear,miniMonth+1,0).getDate();
  const prevLast=new Date(miniYear,miniMonth,0).getDate();

  // 날짜별 이벤트 집계
  const dateMap={};
  merged.filter(r=>r.src===src).forEach(r=>{
    const st=r.status;const isActive=(st==='정상'||st==='유지'||st==='진행중'||st==='');
    const n=diff(r.expDate);
    if(r.expDate&&r.expDate.getFullYear()===miniYear&&r.expDate.getMonth()===miniMonth&&isActive){
      const k=r.expDate.getDate();
      if(!dateMap[k])dateMap[k]={expire:0,suspend:0};
      if(n!==null&&n>=0&&n<=10)dateMap[k].suspend++;
      else dateMap[k].expire++;
    }
    const isRevoke=(st==='정지'||st==='정지(Suspended)');
    const fsscRevoke=r.src==='FSSC'&&(r.fsscStatus||'').toLowerCase()==='suspended';
    if(isRevoke&&(r.src==='IGC'||fsscRevoke)&&r.suspendDate){
      if(r.suspendDate.getFullYear()===miniYear&&r.suspendDate.getMonth()===miniMonth){
        const k=r.suspendDate.getDate();
        if(!dateMap[k])dateMap[k]={expire:0,suspend:0};
        dateMap[k].suspend++;
      }
    }
  });

  // 일정도 포함
  const schedules=getSchedules();
  schedules.forEach(s=>{
    const d=parseDate(s.date);
    if(d&&d.getFullYear()===miniYear&&d.getMonth()===miniMonth){
      const k=d.getDate();
      if(!dateMap[k])dateMap[k]={expire:0,suspend:0};
      if(!dateMap[k].schedule)dateMap[k].schedule=0;
      dateMap[k].schedule++;
    }
  });

  const days=['일','월','화','수','목','금','토'];
  const dCls=['sun','','','','','','sat'];
  let cells=days.map((d,i)=>`<div class="mini-head ${dCls[i]}">${d}</div>`).join('');
  for(let i=0;i<firstDay;i++)cells+=`<div class="mini-cell other-m">${prevLast-firstDay+1+i}</div>`;
  for(let d=1;d<=lastDate;d++){
    const isToday=d===TODAY.getDate()&&miniMonth===TODAY.getMonth()&&miniYear===TODAY.getFullYear();
    const ev=dateMap[d];
    const hasEv=ev&&(ev.expire||ev.suspend||ev.schedule);
    const dots=ev?[ev.suspend?`<span style="background:#ff6b6b"></span>`:'',ev.expire?`<span style="background:#f0a040"></span>`:'',ev.schedule?`<span style="background:#b06aff"></span>`:''].join(''):'';
    cells+=`<div class="mini-cell${isToday?' today-mini':''}${hasEv?' has-event':''}" onclick="showMiniPopup(${miniYear},${miniMonth+1},${d},event)">${d}${hasEv?`<div class="mini-dot">${dots}</div>`:''}</div>`;
  }
  const remain=(7-(firstDay+lastDate)%7)%7;
  for(let i=1;i<=remain;i++)cells+=`<div class="mini-cell other-m">${i}</div>`;

  return`<div class="mini-cal">
    <div class="mini-cal-hd">
      <span>📅 ${miniYear}년 ${mN[miniMonth]}</span>
      <div class="mini-cal-nav">
        <button onclick="miniCalMove(-1)">‹</button>
        <button onclick="miniCalMove(0)">오늘</button>
        <button onclick="miniCalMove(1)">›</button>
      </div>
    </div>
    <div class="mini-grid">${cells}</div>
  </div>`;
}


export function miniCalMove(dir){
  if(dir===0){miniYear=TODAY.getFullYear();miniMonth=TODAY.getMonth();}
  else{miniMonth+=dir;if(miniMonth>11){miniMonth=0;miniYear++;}else if(miniMonth<0){miniMonth=11;miniYear--;}}
  requestDashRefresh();
}


export function showMiniPopup(y,m,d,e){
  e.stopPropagation();
  const src=getMiniSrc();
  const items=[];
  merged.filter(r=>r.src===src).forEach(r=>{
    const st=r.status;const isActive=(st==='정상'||st==='유지'||st==='진행중'||st==='');
    const n=diff(r.expDate);
    if(r.expDate&&r.expDate.getFullYear()===y&&r.expDate.getMonth()===m-1&&r.expDate.getDate()===d&&isActive)
      items.push({label:`${n!==null&&n>=0&&n<=10?'🔴 정지예정':'🟠 만료예정'} ${r.nameKo||r.nameEn}`,idx:r.idx});
    const isRevoke=(st==='정지'||st==='정지(Suspended)');
    const fsscRevoke=r.src==='FSSC'&&(r.fsscStatus||'').toLowerCase()==='suspended';
    if(isRevoke&&(r.src==='IGC'||fsscRevoke)&&r.suspendDate&&r.suspendDate.getFullYear()===y&&r.suspendDate.getMonth()===m-1&&r.suspendDate.getDate()===d)
      items.push({label:`🔴 철회예정 ${r.nameKo||r.nameEn}`,idx:r.idx});
  });
  getSchedules().forEach(s=>{
    const sd=parseDate(s.date);
    if(sd&&sd.getFullYear()===y&&sd.getMonth()===m-1&&sd.getDate()===d)
      items.push({label:`🟣 ${s.memo}`,idx:null});
  });
  if(!items.length)return;
  const popup=document.getElementById('cal-popup');
  setTxt('cal-popup-title',`${y}년 ${m}월 ${d}일`);
  document.getElementById('cal-popup-list').innerHTML=items.map(it=>
    it.idx!==null?`<div class="cal-popup-row" onclick="openModal(${it.idx})">${it.label}</div>`
    :`<div class="cal-popup-row" style="cursor:default">${it.label}</div>`
  ).join('');
  popup.style.left=Math.min(e.pageX,window.innerWidth-340)+'px';
  popup.style.top=(e.pageY+10)+'px';
  popup.classList.add('on');
}


// ── 일정 관리 ──

export function getSchedules(){ return fsGet('schedules')||[]; }

export function saveSchedules(arr){ fsSet('schedules',arr); }


export function renderScheduleHtml(){
  const schedules=getSchedules().sort((a,b)=>new Date(a.date)-new Date(b.date));
  const items=schedules.map((s,i)=>`
    <div class="schedule-item">
      <span class="schedule-date">${s.date||''}</span>
      <span class="schedule-text">${s.memo||''}</span>
      <button class="schedule-del" onclick="delSchedule(${i})">×</button>
    </div>`).join('');
  return`<div class="schedule-wrap">
    <div class="schedule-hd">📋 일정 / 메모 <span class="badge" style="font-size:10px">${schedules.length}</span></div>
    <div class="schedule-list">${items||'<div style="padding:12px 14px;font-size:12px;color:var(--muted)">일정이 없어요</div>'}</div>
    <div class="schedule-add">
      <input type="date" id="sc-date" value="${fmt(TODAY)}">
      <input type="text" id="sc-memo" placeholder="일정 또는 메모 입력..." onkeydown="if(event.key==='Enter')addSchedule()">
      <button class="tbtn pri" style="padding:5px 10px" onclick="addSchedule()">+</button>
    </div>
  </div>`;
}


// ── 상태 세터 (다른 모듈은 직접 재할당 대신 이 함수를 통해 상태를 변경) ──
export function setShowHidden(v){ showHidden=v; }
export function setCurSubTab(v){ curSubTab=v; }

// ── 화면 갱신 요청 (하위 모듈 → 상위(app.js) 순환참조 없이 재렌더 요청) ──
export function requestDashRefresh(){ document.dispatchEvent(new Event('app:refresh-dash')); }
export function requestFullRefresh(){ document.dispatchEvent(new Event('app:refresh-all')); }

// ── 저장된 원본 업로드 데이터 복원 (앱 시작 시 호출) ──
export function hydrateRawData(){
  const fsscData=fsGet('fsscRaw'),fsscName=fsGet('fsscName');
  if(fsscData&&fsscData.length&&'Client number' in fsscData[0]){fsscRaw=fsscData;if(fsscName)setTxt('lbl-fssc',fsscName);}
  const jobData=fsGet('jobRaw'),jobName=fsGet('jobName');
  if(jobData&&jobData.length&&'New Job No.' in jobData[0]){jobRaw=jobData;if(jobName)setTxt('lbl-job',jobName);}
  const igcName=fsGet('igcName');if(igcName)setTxt('lbl-igc',igcName);
}


// onclick="..." 문자열에서 참조하는 이 파일의 함수들을 전역에 노출
Object.assign(window, {loadFSSC, loadIGC, loadJob, mergeAndRender, miniCalMove, showMiniPopup});
