/**
 * ============================================================
 * © 2026 GEG화성 (깊이 e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」에 보호받는 저작물입니다.
 * - 복제권(제16조)·공중송신권(제18조)·배포권(제20조)은
 *   저작권자에게 있습니다.
 * - 정상 경로로 받은 이용자라도 코드의 무단 복제·재배포·
 *   재판매·리브랜딩은 허용되지 않습니다.
 * - 무단 이용 시 「저작권법」 제136조(5년 이하 징역 또는
 *   5천만 원 이하 벌금) 및 제125조(손해배상) 적용 대상이
 *   될 수 있습니다.
 * - 이용 문의: bacusiki777@gmail.com, for2102@jimj.kr
 * ============================================================
 */

// 빌드 서명
const _BUILD_SIG = 'GEGHS-DEEPE-2026';

// 출처 확인용 함수
function getBuildInfo() {
  return {
    sig: _BUILD_SIG,
    owner: 'GEG화성 (깊이 e끌림)',
    year: 2026
  };
}

/*************************************************************
 * 🌍 교실 무역 게임 — 우리 반 전용 시트 백엔드
 *
 *  ▸ 이 시트(=사본 1개)가 한 선생님(한 반)의 데이터를 담습니다.
 *  ▸ 메뉴 [🌍 교실 무역 게임 → 시트 처음 설정하기] 를 누르면
 *    필요한 탭이 자동으로 만들어집니다.
 *  ▸ [배포 → 새 배포 → 웹 앱]으로 배포하면 .../exec 주소가 생기고,
 *    그 주소를 학생 앱 ⚙️ 에 붙여넣으면 이 시트에서만 읽고 씁니다.
 *  ▸ 브라우저에서 바로 부르기 위해 JSONP(GET + callback) 방식을 씁니다.
 *
 *  자세한 순서는 시트의 [선생님 가이드] 탭을 참고하세요.
 *************************************************************/

/** 탭 이름 (모두 한글) */
var TAB = {
  GUIDE:  '사용 설명',
  ROSTER: '학생 명단',
  LOG:    '활동 기록',
  CONFIG: '게임 설정'
};

/** 명단/기록에서 "안내용 임시 텍스트"로 보고 자동으로 걸러낼 패턴 */
var PLACEHOLDER = /(^예\))|예시|샘플|보기\)|여기에|입력하세요|적으세요|바꿔\s*쓰|지우|^이름$|^번호$|←|→|※/;

function isPlaceholder_(v) {
  if (v === null || v === undefined) return true;
  var t = String(v).trim();
  if (!t) return true;
  return PLACEHOLDER.test(t);
}

/* =========================================================
 *  웹 앱 진입점 (JSONP)
 * ========================================================= */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback || 'callback';
  var payload;
  try {
    payload = route_(p);
  } catch (err) {
    payload = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(payload) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function route_(p) {
  switch (p.action) {
    case 'ping':     return ping_();
    case 'roster':   return roster_();
    case 'settings': return settings_();
    case 'teams':    return teams_();
    case 'recent':   return recent_(Number(p.limit) || 20);
    case 'log':      return log_(p);
    default:         return { ok: false, error: '알 수 없는 요청입니다.' };
  }
}

/* =========================================================
 *  읽기
 * ========================================================= */
function ping_() {
  var s = settings_();
  return { ok: true, className: s.className, time: new Date().toISOString() };
}

/** 학생 명단 — 안내용 임시행은 걸러서 돌려줌 */
function roster_() {
  var sh = sheet_(TAB.ROSTER);
  var last = sh.getLastRow();
  var out = [];
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, 2).getValues(); // A:번호, B:이름
    rows.forEach(function (r) {
      var no = r[0], name = r[1];
      if (isPlaceholder_(name)) return;      // 빈 칸 / 안내 텍스트 제외
      out.push({ no: (no === '' ? '' : no), name: String(name).trim() });
    });
  }
  return { ok: true, students: out };
}

/** 게임 설정 — 키/값 표를 읽어 숫자로 변환 (없으면 기본값) */
function settings_() {
  var sh = sheet_(TAB.CONFIG);
  var last = sh.getLastRow();
  var cfg = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      var k = String(r[0]).trim();
      if (k) cfg[k] = r[1];
    });
  }
  return {
    ok: true,
    className:  String(cfg['반 이름'] || '').trim(),
    rounds:     Number(cfg['라운드 수']) || 10,
    seconds:    Number(cfg['라운드 시간(초)']) || 60,
    startMoney: Number(cfg['시작 자금']) || 300
  };
}

/** 팀 편성 — [학생 명단]의 이름 옆 'C열(팀)'에서 읽음. 팀이 비면 제외(1인 플레이) */
function teams_() {
  var sh = sheet_(TAB.ROSTER);
  var last = sh.getLastRow();
  var out = [];
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, 3).getValues(); // A:번호, B:이름, C:팀
    rows.forEach(function (r) {
      var name = String(r[1] || '').trim();
      var team = String(r[2] || '').trim();
      if (isPlaceholder_(name)) return;
      if (!team) return;                 // 팀 칸이 비면 팀 없음(컴퓨터와 1인 플레이)
      out.push({ team: team, name: name, country: '' }); // 나라는 앱이 자동 배정
    });
  }
  return { ok: true, teams: out };
}

/** 최근 활동 — 안내용 임시행은 걸러서 돌려줌 (앱에서 필요 시 사용) */
function recent_(limit) {
  var sh = sheet_(TAB.LOG);
  var last = sh.getLastRow();
  var out = [];
  if (last >= 2) {
    var n = Math.min(last - 1, 200);
    var rows = sh.getRange(last - n + 1, 1, n, 6).getValues(); // 시각,이름,종류,내용,점수,등수
    for (var i = rows.length - 1; i >= 0 && out.length < limit; i--) {
      var r = rows[i];
      if (isPlaceholder_(r[1])) continue;
      out.push({
        time:   fmtTime_(r[0]),
        name:   String(r[1]).trim(),
        kind:   String(r[2] || '').trim(),
        detail: String(r[3] || '').trim(),
        score:  r[4], rank: r[5]
      });
    }
  }
  return { ok: true, events: out };
}

/* =========================================================
 *  쓰기 (학생 활동 기록)
 * ========================================================= */
function log_(p) {
  var name = String(p.name || '').trim();
  if (isPlaceholder_(name)) return { ok: false, error: '이름이 없습니다.' };

  var kind   = String(p.kind || '').trim();
  var detail = String(p.detail || '').trim();
  var score  = (p.score  != null && p.score  !== '') ? Number(p.score) : '';
  var rank   = (p.rank   != null && p.rank   !== '') ? Number(p.rank)  : '';
  var now = new Date();

  // 1) [활동 기록] 탭에 한 줄 추가 (전체 보관)
  sheet_(TAB.LOG).appendRow([now, name, kind, detail, score, rank]);

  // 2) [학생 명단] 탭에서 이름이 같은 행을 찾아 '최근 활동'을 이름 옆(같은 줄)에 기록
  rosterLatest_(name, kind, detail, score, now);

  return { ok: true };
}

/** 학생 명단에서 이름이 일치하는 행의 C~E(최근 활동·점수·갱신 시각)를 갱신 */
function rosterLatest_(name, kind, detail, score, when) {
  var sh = sheet_(TAB.ROSTER);
  var last = sh.getLastRow();
  if (last < 2) return;
  var names = sh.getRange(2, 2, last - 1, 1).getValues(); // B열(이름)
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === name) {
      var row = i + 2;
      var text = (kind ? '[' + kind + '] ' : '') + detail;
      sh.getRange(row, 4, 1, 3).setValues([[text, (score === '' ? '' : score), fmtTime_(when)]]); // D,E,F
      return;
    }
  }
  // 명단에 없는 이름이면 기록하지 않음 (명단의 학생만 참여)
}

/* =========================================================
 *  결과 보기 — 사이드바 & 대시보드
 * ========================================================= */
function showResultsSidebar() {
  var html = HtmlService.createHtmlOutput(resultsHtml_('sidebar')).setTitle('📋 결과 보기');
  SpreadsheetApp.getUi().showSidebar(html);
}
function showResultsDashboard() {
  var html = HtmlService.createHtmlOutput(resultsHtml_('dashboard')).setWidth(920).setHeight(660);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 결과 대시보드');
}

/** 사이드바/대시보드가 불러가는 결과 데이터 (google.script.run으로 호출) */
function dashboardData() {
  var ss = SpreadsheetApp.getActive();
  var rs = ss.getSheetByName(TAB.ROSTER);
  var students = [];
  if (rs) {
    var last = rs.getLastRow();
    if (last >= 2) {
      var v = rs.getRange(2, 1, last - 1, 6).getValues(); // 번호,이름,팀,최근활동,점수,시각
      v.forEach(function (r) {
        var name = String(r[1] || '').trim();
        if (isPlaceholder_(name)) return;
        var act = String(r[3] || '').trim();
        var score = (r[4] === '' || r[4] == null) ? null : Number(r[4]);
        students.push({
          no: r[0], name: name, team: String(r[2] || '').trim(),
          activity: act, score: score, time: String(r[5] || '').trim(),
          done: /종료/.test(act)
        });
      });
    }
  }
  var ranked = students.filter(function (s) { return s.score != null; })
                       .sort(function (a, b) { return b.score - a.score; });
  var noScore = students.filter(function (s) { return s.score == null; });

  var tmap = {};
  students.forEach(function (s) {
    if (!s.team) return;
    if (!tmap[s.team]) tmap[s.team] = { team: s.team, sum: 0, cnt: 0, max: 0 };
    if (s.score != null) { tmap[s.team].sum += s.score; tmap[s.team].cnt++; if (s.score > tmap[s.team].max) tmap[s.team].max = s.score; }
  });
  var teams = Object.keys(tmap).map(function (k) {
    var t = tmap[k];
    return { team: t.team, sum: t.sum, cnt: t.cnt, max: t.max, avg: t.cnt ? Math.round(t.sum / t.cnt) : 0 };
  }).sort(function (a, b) { return b.sum - a.sum; });

  return {
    ranked: ranked, noScore: noScore, teams: teams,
    king: ranked.length ? ranked[0] : null,
    stats: { total: students.length, played: ranked.length, done: students.filter(function (s) { return s.done; }).length },
    updated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm:ss')
  };
}

/** 사이드바/대시보드 공용 HTML (mode: 'sidebar' | 'dashboard') */
function resultsHtml_(mode) {
  var isDash = (mode === 'dashboard');
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_top"><style>
  body{font-family:"Apple SD Gothic Neo","Malgun Gothic",sans-serif;margin:0;padding:12px;color:#1f2430;background:#f7f8fa;}
  .top{display:flex;align-items:center;justify-content:space-between;gap:8px;}
  h2{margin:0;font-size:${isDash ? '20px' : '15px'};}
  .refresh{border:none;border-radius:8px;background:#4f9cf9;color:#fff;font-weight:700;padding:7px 12px;cursor:pointer;font-family:inherit;font-size:13px;}
  .updated{color:#8b90a0;font-size:12px;margin:6px 0 10px;}
  .king{background:linear-gradient(135deg,#fff3cd,#ffe69a);border:2px solid #f5c542;border-radius:12px;padding:11px 13px;margin-bottom:12px;font-weight:800;font-size:${isDash ? '17px' : '14px'};}
  .king .s{color:#9a6b00;}
  .sec{font-weight:800;margin:16px 0 6px;color:#394052;font-size:13px;}
  .row{display:flex;align-items:center;gap:8px;padding:7px 6px;border-bottom:1px solid #e7e9ee;font-size:14px;}
  .pos{width:26px;text-align:center;font-weight:800;color:#6b7280;flex:none;}
  .nm{flex:1;font-weight:700;min-width:0;}
  .tm{font-size:11px;color:#5a6477;background:#eef1f6;border-radius:6px;padding:1px 6px;margin-left:4px;}
  .done{font-size:11px;color:#2e9b63;margin-left:4px;}
  .sc{font-weight:800;color:#1a7f4b;flex:none;}
  .bar{height:7px;background:#e7e9ee;border-radius:5px;overflow:hidden;margin-top:4px;}
  .bar>i{display:block;height:100%;background:linear-gradient(90deg,#4f9cf9,#36c275);}
  .muted{color:#8b90a0;font-size:12px;}
  .stats{margin:2px 0 4px;}
  .stat{display:inline-block;background:#fff;border:1px solid #e7e9ee;border-radius:10px;padding:6px 10px;margin:0 6px 6px 0;font-size:12px;}
  .stat b{font-size:15px;color:#1f2430;}
  .grid{display:${isDash ? 'grid' : 'block'};grid-template-columns:1fr 1fr;gap:18px;}
  </style></head><body>
  <div class="top"><h2>${isDash ? '📊 결과 대시보드' : '📋 결과'}</h2><button class="refresh" onclick="load()">↻ 새로고침</button></div>
  <div id="updated" class="updated">불러오는 중…</div>
  <div id="content"></div>
  <script>
  function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;');}
  function rowHtml(i,s,max){
    var w = (s.score!=null && max) ? Math.round(s.score/max*100) : 0;
    return '<div class="row"><span class="pos">'+(i==null?'-':(i+1))+'</span>'+
      '<span class="nm">'+esc(s.name)+(s.team?'<span class="tm">'+esc(s.team)+'</span>':'')+(s.done?'<span class="done">완료</span>':'')+
      (s.score!=null?'<div class="bar"><i style="width:'+w+'%"></i></div>':'')+'</span>'+
      (s.score!=null?'<span class="sc">'+s.score+'</span>':'<span class="muted">'+(esc(s.activity)||'미시작')+'</span>')+'</div>';
  }
  function render(d){
    document.getElementById('updated').innerHTML='<span class="stats"><span class="stat">참여 <b>'+d.stats.played+'</b>/'+d.stats.total+'명</span><span class="stat">완료 <b>'+d.stats.done+'</b>명</span></span><br>갱신: '+esc(d.updated);
    var max = d.ranked.length ? d.ranked[0].score : 1; if(!max) max=1;
    var h='';
    if(d.king){ h+='<div class="king">👑 무역왕 &nbsp;'+esc(d.king.name)+(d.king.team?' ('+esc(d.king.team)+')':'')+' &nbsp;<span class="s">점수 '+d.king.score+'</span></div>'; }
    var left='<div><div class="sec">학생 순위</div>';
    if(d.ranked.length){ d.ranked.forEach(function(s,i){ left+=rowHtml(i,s,max); }); }
    else { left+='<p class="muted">아직 점수가 기록된 학생이 없어요.</p>'; }
    if(d.noScore.length){ left+='<div class="sec">아직 시작 안 함 / 점수 없음</div>'; d.noScore.forEach(function(s){ left+=rowHtml(null,s,max); }); }
    left+='</div>';
    var right='';
    ${isDash ? `if(d.teams.length){ right='<div><div class="sec">팀 순위 (점수 합계)</div>'; var tmax=d.teams[0].sum||1; d.teams.forEach(function(t,i){ var w=Math.round(t.sum/tmax*100); right+='<div class="row"><span class="pos">'+(i+1)+'</span><span class="nm">'+esc(t.team)+' <span class="muted">'+t.cnt+'명·평균 '+t.avg+'</span><div class="bar"><i style="width:'+w+'%"></i></div></span><span class="sc">'+t.sum+'</span></div>'; }); right+='</div>'; }` : ''}
    document.getElementById('content').innerHTML = h + '<div class="grid">'+left+right+'</div>';
  }
  function load(){
    document.getElementById('updated').textContent='불러오는 중…';
    google.script.run.withSuccessHandler(render).withFailureHandler(function(e){ document.getElementById('content').innerHTML='<p class="muted">불러오지 못했어요: '+esc(e&&e.message)+'</p>'; }).dashboardData();
  }
  load();
  </script></body></html>`;
}

/* =========================================================
 *  시트 메뉴 & 초기 설정
 * ========================================================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌍 교실 무역 게임')
    .addItem('📋 시트 처음 설정하기', 'setupSheets')
    .addItem('📖 사용 설명 보기', 'showGuide')
    .addSeparator()
    .addItem('📊 결과 대시보드', 'showResultsDashboard')
    .addItem('📋 결과 사이드바', 'showResultsSidebar')
    .addSeparator()
    .addItem('🧹 활동 기록 비우기', 'clearLog')
    .addToUi();
}

function setupSheets() {
  var ss = SpreadsheetApp.getActive();
  ensureGuide_(ss);
  ensureRoster_(ss);
  ensureLog_(ss);
  ensureConfig_(ss);
  cleanupDefault_(ss);
  ss.setActiveSheet(ss.getSheetByName(TAB.ROSTER));
  ss.toast('준비 완료! [학생 명단]에서 학생1~학생25를 실제 이름으로 바꾸세요.', '🌍 교실 무역 게임', 6);
}

function showGuide() {
  var ss = SpreadsheetApp.getActive();
  ensureGuide_(ss);
  ss.setActiveSheet(ss.getSheetByName(TAB.GUIDE));
  ss.toast('설정·팀 편성·연결 순서가 이 탭에 정리되어 있어요.', '📖 사용 설명', 6);
}

function clearLog() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('활동 기록 비우기',
    '지금까지 쌓인 활동 기록을 모두 지울까요? (명단은 그대로 유지됩니다)',
    ui.ButtonSet.OK_CANCEL);
  if (res !== ui.Button.OK) return;

  var log = sheet_(TAB.LOG);
  if (log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 6).clearContent();
  }
  var roster = sheet_(TAB.ROSTER);
  if (roster.getLastRow() > 1) {
    roster.getRange(2, 4, roster.getLastRow() - 1, 3).clearContent(); // D~F: 학생별 최근 활동 비우기 (팀 칸 C는 유지)
  }
  SpreadsheetApp.getActive().toast('활동 기록을 비웠어요.', '🧹', 4);
}

/* ---------- 탭 생성/정리 helpers ---------- */
function sheet_(name) {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureGuide_(ss) {
  // 옛 이름(선생님 가이드)이나 같은 이름의 탭이 있으면 지우고 새로 만든다 (내용 재사용 안 함)
  ['선생님 가이드', TAB.GUIDE].forEach(function (nm) {
    var old = ss.getSheetByName(nm);
    if (old) ss.deleteSheet(old);
  });
  var sh = ss.insertSheet(TAB.GUIDE, 0); // 첫 번째 위치

  // [내용, 종류]  — 종류 sec 은 항목 순서대로 번호를 자동으로 붙임
  var rows = [
    ['교실 무역 게임 - 사용 설명', 'title'],
    ['', 'blank'],
    ['데이터는 설정을 바꿀 때도 앱 화면이 아니라 해당 시트 탭에서 직접 수정하세요. 탭 이름은 코드에 연결되어 있으므로 삭제하거나 변경하지 마세요.', 'body'],
    ['이 시트는 우리 반 전용입니다. 다른 반이나 다른 선생님과 데이터가 섞이지 않습니다.', 'body'],
    ['', 'blank'],

    ['시트 처음 설정하기', 'sec'],
    ['위 메뉴에서 교실 무역 게임, 시트 처음 설정하기를 누르면 학생 명단, 활동 기록, 게임 설정 탭이 자동으로 만들어집니다.', 'body'],
    ['학생 명단 탭에서 학생1부터 학생25까지를 실제 이름으로 바꿉니다. 학생이 더 많으면 아래에 줄을 추가합니다.', 'body'],
    ['', 'blank'],

    ['팀 편성 (선택)', 'sec'],
    ['학생 명단의 이름 옆 팀 칸에 같은 팀끼리 같은 이름을 적습니다. 예를 들어 1모둠이라고 적습니다.', 'body'],
    ['한 팀은 최대 여섯 명까지입니다. 나라가 여섯 개라서 일곱 번째 학생부터는 맡을 자리가 없습니다.', 'body'],
    ['한 팀에 여섯 명이 안 되면 비어 있는 나라 자리는 컴퓨터가 대신 맡습니다.', 'body'],
    ['팀 칸을 비우면 그 학생은 컴퓨터와 혼자 플레이합니다.', 'body'],
    ['', 'blank'],

    ['연결 주소 만들기 (한 번만)', 'sec'],
    ['확장 프로그램 메뉴에서 Apps Script를 엽니다.', 'body'],
    ['오른쪽 위에서 배포, 새 배포를 누르고 유형을 웹 앱으로 고릅니다.', 'body'],
    ['액세스 권한을 모든 사용자로 설정합니다. 이 설정을 하지 않으면 학생이 연결되지 않습니다.', 'body'],
    ['배포를 누르고 권한 승인 창이 나오면 허용한 뒤 만들어진 주소를 복사합니다. 주소는 exec로 끝납니다.', 'body'],
    ['', 'blank'],

    ['학생 입장', 'sec'],
    ['앱 첫 화면의 선생님 메뉴에 위 주소를 붙여넣고 연결하기를 누릅니다.', 'body'],
    ['그러면 학생 공유 링크가 만들어집니다. 이 링크를 학생에게 주면 학생 화면에서는 설정이 보이지 않고 바로 명단이 연결됩니다.', 'body'],
    ['명단에 있는 이름만 입장할 수 있습니다.', 'body'],
    ['', 'blank'],

    ['진행 기록 보기', 'sec'],
    ['학생이 활동하면 그 학생 이름 옆 최근 활동이 자동으로 갱신됩니다.', 'body'],
    ['전체 흐름은 활동 기록 탭에 쌓입니다. 지우려면 메뉴의 활동 기록 비우기를 사용합니다.', 'body'],
    ['결과는 교실 무역 게임 메뉴의 결과 대시보드와 결과 사이드바에서도 볼 수 있습니다.', 'body'],
    ['', 'blank'],

    ['연결하지 않으면 (체험 모드)', 'sec'],
    ['시트를 연결하지 않아도 학생1부터 학생30까지 골라 바로 체험할 수 있습니다. 이때 진행은 그 기기에만 저장되며 한 판을 끝내면 저장이 비워집니다.', 'body'],
    ['', 'blank'],

    ['코드를 고쳐 다시 배포할 때는 배포, 배포 관리, 수정, 새 버전을 차례로 고릅니다. 주소는 그대로 유지됩니다.', 'body'],
    ['문의: bacusiki777@gmail.com, for2102@jimj.kr', 'body']
  ];

  // 섹션 번호를 항목 순서대로 자동으로 매겨 한 번에 입력
  var n = 0;
  var content = rows.map(function (r) {
    var t = r[0];
    if (r[1] === 'sec') { n++; t = n + '. ' + t; }
    return [t];
  });
  sh.getRange(1, 1, content.length, 1).setValues(content);

  sh.setColumnWidth(1, 760);
  var all = sh.getRange(1, 1, rows.length, 1);
  all.setWrap(true).setVerticalAlignment('middle').setFontSize(11).setFontColor('#2a2f3a')
     .setBorder(true, true, true, true, true, true, '#d7dde6', SpreadsheetApp.BorderStyle.SOLID);

  for (var i = 0; i < rows.length; i++) {
    var tag = rows[i][1];
    var cell = sh.getRange(i + 1, 1);
    if (tag === 'title') cell.setBackground('#dfe8f5').setFontWeight('bold').setFontSize(14);
    else if (tag === 'sec') cell.setBackground('#eef3fa').setFontWeight('bold').setFontSize(13);
    else cell.setBackground('#ffffff');
  }
  sh.setFrozenRows(1);
  sh.setHiddenGridlines(true);
  return sh;
}

function ensureRoster_(ss) {
  var sh = ss.getSheetByName(TAB.ROSTER) || ss.insertSheet(TAB.ROSTER);
  // 선생님이 채우는 칸(A:번호 B:이름 C:팀)  +  앱이 자동으로 채우는 칸(D:최근 활동 E:점수 F:갱신 시각)
  sh.getRange('A1:F1').setValues([['번호', '이름', '팀', '최근 활동', '점수', '갱신 시각']]);
  sh.getRange('A1:C1').setFontWeight('bold').setBackground('#1b2a41').setFontColor('#ffffff');
  sh.getRange('D1:F1').setFontWeight('bold').setBackground('#274a7d').setFontColor('#ffffff');
  sh.setFrozenRows(1);

  // 학생1~학생25 미리 채우기 (이미 입력돼 있으면 덮어쓰지 않음)
  if (isPlaceholder_(sh.getRange('B2').getValue())) {
    var seed = [];
    for (var i = 1; i <= 25; i++) seed.push([i, '학생' + i]);
    sh.getRange(2, 1, 25, 2).setValues(seed);
  }
  // 안내는 데이터가 아닌 '메모'로 — 앱 읽기에 영향 없음
  sh.getRange('B1').setNote('여기 학생1~학생25를 실제 이름으로 바꿔 주세요.\n학생이 더 많으면 아래에 줄을 추가하세요.\n안내용 글자(예: "예)", "여기에 입력")는 앱에 표시되지 않습니다.');
  sh.getRange('C1').setNote('같은 팀끼리 같은 팀 이름을 적으세요. 예) 1모둠, 1모둠, 2모둠 …\n비워두면 그 학생은 컴퓨터와 1인 플레이를 합니다.\n같은 팀의 빈 나라 자리는 컴퓨터가 맡아요.');
  sh.getRange('D1').setNote('학생이 활동하면 앱이 자동으로 채우는 칸이에요. (직접 쓰지 않아도 됩니다)');

  sh.setColumnWidth(1, 56);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 300);
  sh.setColumnWidth(5, 70);
  sh.setColumnWidth(6, 130);
  return sh;
}

function ensureLog_(ss) {
  var sh = ss.getSheetByName(TAB.LOG) || ss.insertSheet(TAB.LOG);
  sh.getRange('A1:F1').setValues([['시각', '이름', '종류', '내용', '점수', '등수']]);
  sh.getRange('A1:F1').setFontWeight('bold').setBackground('#1b2a41').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  // 예시 한 줄(자동으로 걸러지는 안내행) — 비어 있을 때만
  if (sh.getLastRow() < 2) {
    sh.appendRow([new Date(), '예) 홍길동', '안내', '이 줄은 예시예요 — 게임 화면에는 표시되지 않아요', '', '']);
    sh.getRange(2, 1, 1, 6).setFontColor('#9aa7bd').setFontStyle('italic');
  }
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 70);
  sh.setColumnWidth(4, 320);
  return sh;
}

function ensureConfig_(ss) {
  var sh = ss.getSheetByName(TAB.CONFIG) || ss.insertSheet(TAB.CONFIG);
  sh.getRange('A1:B1').setValues([['설정 항목', '값']]);
  sh.getRange('A1:B1').setFontWeight('bold').setBackground('#1b2a41').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 4, 2).setValues([
      ['반 이름', ''],
      ['라운드 수', 10],
      ['라운드 시간(초)', 60],
      ['시작 자금', 300]
    ]);
    sh.getRange('B2').setNote('비워두면 학생 화면에 "우리 반"으로 표시돼요. 예: 3학년 2반');
  }
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 140);
  return sh;
}

/** 새 스프레드시트에 남아있는 빈 기본 탭(시트1/Sheet1) 정리 */
function cleanupDefault_(ss) {
  ['시트1', 'Sheet1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch (e) {}
    }
  });
}

function fmtTime_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd HH:mm');
}
