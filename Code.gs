/*******************************************************************
 * 우면동교회 출석부 - 다부서 백엔드 (Google Apps Script)
 *
 * ★ 기존에 쓰던 스크립트를 이 코드로 교체한 뒤,
 *   "배포 > 배포 관리 > (연필) 편집 > 버전: 새 버전 > 배포" 로 갱신하세요.
 *   (새 배포가 아니라 "새 버전"으로 하면 웹앱 주소가 그대로 유지됩니다.)
 *
 * - 부서(유아부·유치부·유년부·초등부·청소년부)별로 데이터를 나눠 저장합니다.
 * - 예전 초등부 단일 데이터는 자동으로 "초등부" 부서로 옮겨 보존됩니다.
 *******************************************************************/

var DATA_SHEET = "_data";
var CHUNK = 45000;
var DEPTS = ["유아부","유치부","유년부","초등부","청소년부"];
var DEFAULT_TALENTS = {attendance:2, qt:1, memorize:3, family:2, transcribe:2};

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function dataSheet(){ var s=ss(); var sh=s.getSheetByName(DATA_SHEET); if(!sh){ sh=s.insertSheet(DATA_SHEET); sh.hideSheet(); } return sh; }

function readRoot(){
  var sh=dataSheet(); var last=sh.getLastRow(); if(last<1) return null;
  var vals=sh.getRange(1,1,last,1).getValues(); var str="";
  for(var i=0;i<vals.length;i++) str+=vals[i][0];
  if(!str) return null;
  try { return migrateToRoot(JSON.parse(str)); } catch(e){ return null; }
}
function writeRoot(obj){
  var sh=dataSheet(); var str=JSON.stringify(obj); sh.clearContents();
  var rows=[]; for(var i=0;i<str.length;i+=CHUNK) rows.push([str.substring(i,i+CHUNK)]);
  if(rows.length) sh.getRange(1,1,rows.length,1).setValues(rows);
}

/* ---------- 웹 요청 ---------- */
function doGet(e){ var r=readRoot(); return json(r || defaultRoot()); }
function doPost(e){
  var lock=LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(err){ return json({ok:false, error:"busy"}); }
  try {
    var body=JSON.parse(e.postData.contents);
    var incoming=(body && body.state!==undefined) ? body.state : body;
    var cur=readRoot();
    var merged=mergeRoot(cur || defaultRoot(), incoming);
    writeRoot(merged);
    renderReadable(merged);
    return json({ok:true});
  } catch(err){ return json({ok:false, error:String(err)}); }
  finally { lock.releaseLock(); }
}
function json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- 구조/마이그레이션 ---------- */
function blankDept(){ return {classes:[], students:[], records:{}, talents:cloneTalents(), seq:1, cfgUpdatedAt:0, deleted:[]}; }
function cloneTalents(){ return {attendance:DEFAULT_TALENTS.attendance, qt:DEFAULT_TALENTS.qt, memorize:DEFAULT_TALENTS.memorize, family:DEFAULT_TALENTS.family, transcribe:DEFAULT_TALENTS.transcribe}; }
function defaultRoot(){ var d={}; for(var i=0;i<DEPTS.length;i++) d[DEPTS[i]]=blankDept(); return {v:2, depts:d}; }
function migrateToRoot(obj){
  if(!obj || typeof obj!=="object") return defaultRoot();
  if(obj.depts){ var r=defaultRoot(); for(var i=0;i<DEPTS.length;i++){ var k=DEPTS[i]; if(obj.depts[k]) r.depts[k]=Object.assign(blankDept(), obj.depts[k]); } return r; }
  var r2=defaultRoot();
  if(obj.students || obj.classes) r2.depts["초등부"]=Object.assign(blankDept(), obj);
  return r2;
}

/* ---------- 병합 ---------- */
function mergeState(local, remote){
  local = local || blankDept();
  if(!remote || typeof remote!=="object") return local;
  var del={};
  (local.deleted||[]).forEach(function(id){ del[id]=1; });
  (remote.deleted||[]).forEach(function(id){ del[id]=1; });
  var remoteNewer=(remote.cfgUpdatedAt||0)>(local.cfgUpdatedAt||0);
  var base=remoteNewer?remote:local;
  var out={ classes:(base.classes||[]).slice(), talents:Object.assign(cloneTalents(), base.talents||{}),
    cfgUpdatedAt:Math.max(local.cfgUpdatedAt||0, remote.cfgUpdatedAt||0), seq:Math.max(local.seq||1, remote.seq||1),
    deleted:Object.keys(del), students:[], records:{} };
  var byId={};
  (local.students||[]).forEach(function(s){ if(!del[s.id]) byId[s.id]=s; });
  (remote.students||[]).forEach(function(s){ if(del[s.id])return; if(!byId[s.id]||remoteNewer) byId[s.id]=s; });
  out.students=Object.keys(byId).map(function(k){ return byId[k]; });
  var sids={};
  Object.keys(local.records||{}).forEach(function(k){ sids[k]=1; });
  Object.keys(remote.records||{}).forEach(function(k){ sids[k]=1; });
  Object.keys(sids).forEach(function(sid){
    if(del[sid]) return;
    var lw=(local.records||{})[sid]||{}, rw=(remote.records||{})[sid]||{};
    var wks={}; Object.keys(lw).forEach(function(w){ wks[w]=1; }); Object.keys(rw).forEach(function(w){ wks[w]=1; });
    var m={};
    Object.keys(wks).forEach(function(wk){ var a=lw[wk], b=rw[wk]; m[wk]=(a&&b)?(((b.u||0)>(a.u||0))?b:a):(a||b); });
    if(Object.keys(m).length) out.records[sid]=m;
  });
  return out;
}
function mergeRoot(local, remote){
  local = local || defaultRoot();
  if(!remote || typeof remote!=="object") return local;
  var rd = remote.depts ? remote.depts : migrateToRoot(remote).depts;
  var out=defaultRoot();
  for(var i=0;i<DEPTS.length;i++){ var d=DEPTS[i]; out.depts[d]=mergeState(local.depts[d]||blankDept(), rd[d]||blankDept()); }
  return out;
}

/* ---------- 달란트 ---------- */
function calcTalent(r, t){
  if(!r) return 0; t=t||DEFAULT_TALENTS; var v=0;
  v+=r.attendance?t.attendance:0; v+=(r.qt||0)*t.qt; v+=r.memorize?t.memorize:0; v+=r.family?t.family:0; v+=r.transcribe?t.transcribe:0;
  return v;
}

/* ---------- 사람이 보기 좋은 시트 (부서 열 포함) ---------- */
function renderReadable(root){
  var mBody=[["부서","반","이름"]];
  var cBody=[["부서","반","이름","학생 연락처","학부모 성함","학부모 연락처"]];
  var rBody=[["부서","주간(주일)","반","이름","출석","QT(0-7)","성경암송","가정예배","성경필사","달란트","메모/기도제목"]];
  for(var i=0;i<DEPTS.length;i++){
    var dep=DEPTS[i]; var st=root.depts[dep]; if(!st) continue; var tal=st.talents||DEFAULT_TALENTS;
    st.students.forEach(function(s){
      mBody.push([dep, s.className||"", s.name||""]);
      cBody.push([dep, s.className||"", s.name||"", s.phone||"", s.parent||"", s.parentPhone||""]);
    });
    var weeks={}; Object.keys(st.records||{}).forEach(function(sid){ Object.keys(st.records[sid]).forEach(function(w){ weeks[w]=1; }); });
    Object.keys(weeks).sort().forEach(function(wk){
      st.students.forEach(function(s){
        var r=st.records[s.id] && st.records[s.id][wk]; if(!r) return;
        rBody.push([dep, wk, s.className||"", s.name||"", r.attendance?"O":"", r.qt||0, r.memorize?"O":"", r.family?"O":"", r.transcribe?"O":"", calcTalent(r,tal), r.memo||""]);
      });
    });
  }
  putSheet("학생명단", mBody);
  putSheet("연락처", cBody);
  putSheet("주간기록", rBody);
}
function putSheet(name, body){
  var s=ss(); var sh=s.getSheetByName(name); if(!sh) sh=s.insertSheet(name);
  sh.clearContents();
  if(body.length){ sh.getRange(1,1,body.length,body[0].length).setValues(body); sh.getRange(1,1,1,body[0].length).setFontWeight("bold"); sh.setFrozenRows(1); }
}

/* ---------- 최초 1회(선택) ---------- */
function setup(){ dataSheet(); renderReadable(readRoot() || defaultRoot()); }
