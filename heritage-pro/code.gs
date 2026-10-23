// ====================================================================================
// [1] 환경 설정 변수
// ====================================================================================

var SHEET_URL = 'https://docs.google.com/spreadsheets/d/1Hd5bMxVTycTRd2SOr8MrhwOFF3dbCAHyNDxiWUgXvsk/edit?gid=0#gid=0';
const GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY';
const FINANCIAL_SHEET_ID = '1fCOseYUJ4VA1Qv2SaaHRu5tSS8tO8ckQybavkuZyRIg';
const FINANCIAL_SHEET_NAME = '출력결과';
const LIFE_INS_FOLDER_ID = '10ycnsheDXd0RRz9wsKI9e95SSlXxlw-I';
const BOARD_FOLDER_ID = '18k_wHGNjALSI7UNN0uqOZx8IpERHUP3A';
const LINK_LIFE_INS = 'https://drive.google.com/drive/folders/1pHq15mKan0qytePmH1CNOS_j7c6l1BsB';
const PHOTO_FOLDER_ID = '1ptmNBSSC_Qy8YQ_cyhwCdssJmCICrKoZ';

// ====================================================================================
// [2] 웹앱 기본 구동 함수
// ====================================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('헤리티지 영업지원 시스템')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function testConnection() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    return "연결 성공: " + ss.getName();
  } catch (e) {
    return "연결 실패: " + e.toString();
  }
}

function formatDate(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) {
    var y = dateVal.getFullYear();
    var m = dateVal.getMonth() + 1;
    var d = dateVal.getDate();
    return y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
  }
  return String(dateVal).substring(0, 10);
}

function normalizeMonth(value) {
  if (!value) return "";
  if (value instanceof Date) {
    var y = value.getFullYear();
    var m = value.getMonth() + 1;
    return y + '-' + (m < 10 ? '0' + m : m);
  }
  var str = String(value).trim();
  if (str.length < 7) return str;
  return str.substring(0, 7);
}

// ====================================================================================
// [3] 사용자 및 영업 데이터 처리 함수
// ====================================================================================

function getUserInfo(month) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    // [직급/지점 변경 반영] 2026-09부터는 '사용자 정보 2609' 탭을 사용한다.
    // (김태건 CM→지점장 승격, 리더스 지점 신설 및 FC 이동 등)
    // month가 넘어오지 않으면 기존 '사용자 정보' 탭을 사용한다.
    var sheetName = '사용자 정보';
    if (month && String(month) >= '2026-09') {
      var newSheet = ss.getSheetByName('사용자 정보 2609');
      if (newSheet) sheetName = '사용자 정보 2609';
    }
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var users = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        users.push({
          name:       String(row[0]).trim(),
          role:       String(row[1]).trim(),
          branch:     String(row[2]).trim(),
          joinDate:   formatDate(row[3]),
          resignDate: formatDate(row[4]),
          recruiter:  String(row[5]).trim(),
          residentId: String(row[6]).trim(),
          grade:      String(row[7]).trim().toUpperCase(),
          phone:      row[8] ? String(row[8]).trim() : "",
          birthday:   row[9] ? String(row[9]).trim() : "",
          email:      row[10] ? String(row[10]).trim() : ""
        });
      }
    }
    return users;
  } catch (e) {
    return [];
  }
}

function getSheetProductData() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('상품');
    if (!sheet) sheet = ss.getSheets()[2];
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length > 0) data.shift();
    return data;
  } catch (e) {
    return [];
  }
}

// [개편] 2026-07 입력분부터 사용하는 상품2607 탭 요율 로드 (열 구성은 '상품' 탭과 동일)
function getSheetProduct2607Data() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('상품2607');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length > 0) data.shift();
    return data;
  } catch (e) {
    return [];
  }
}

// [상품전략] 구글드라이브 '상품전략' 폴더의 PDF 목록을 반환
//  - 폴더 내 PDF 파일들을 이름순으로 정렬해 [{id, name}] 배열로 반환
//  - iframe 미리보기는 https://drive.google.com/file/d/<id>/preview 형태로 사용
var PRODUCT_STRATEGY_FOLDER_ID = '1F4oZvj3bm5RqQSvcaKIsVkLrUq5cH_b_';
function getProductStrategyFiles() {
  try {
    var folder = DriveApp.getFolderById(PRODUCT_STRATEGY_FOLDER_ID);
    var it = folder.getFilesByType(MimeType.PDF);
    var list = [];
    while (it.hasNext()) {
      var f = it.next();
      list.push({ id: f.getId(), name: f.getName() });
    }
    // 파일명 오름차순 정렬
    list.sort(function(a, b) { return a.name.localeCompare(b.name, 'ko'); });
    return list;
  } catch (e) {
    return [];
  }
}


// ===== [지사목표] '지사목표' 탭에서 헤리티지 목표금액 로드 =====
// 시트 구조: A1='지사명', B1='목표금액'  → A열에서 '헤리티지'가 있는 행의 B열 숫자가 목표금액.
// 별도 문서(BRANCH_GOAL_SHEET_ID)를 먼저 보고, 없으면 앱 데이터 시트(SHEET_URL)에서도 찾는다.
// 실패 시 클라이언트가 기본값(1,200만)으로 되돌린다.
var BRANCH_GOAL_SHEET_ID = '1LejP4uKQoYBmTkjgYObXBaPzE29FJdMgONmcyO-3YA4';
var BRANCH_GOAL_TAB = '지사목표';
var BRANCH_GOAL_NAME = '헤리티지';

function _readBranchGoalFromSpreadsheet(ss, sourceLabel) {
  var out = { ok: false, map: {}, single: null, how: '', reason: '', source: sourceLabel, branch: '' };
  var sheet = ss.getSheetByName(BRANCH_GOAL_TAB);
  if (!sheet) { out.reason = "'" + BRANCH_GOAL_TAB + "' 탭 없음"; return out; }
  var data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) { out.reason = '탭이 비어 있음'; return out; }

  var norm = function(v) { return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ''); };
  var num = function(v) {
    if (typeof v === 'number') return v;
    var s = String(v === null || v === undefined ? '' : v).replace(/[,\s원]/g, '');
    if (s === '') return null;
    var f = parseFloat(s);
    return isNaN(f) ? null : f;
  };

  // 1) 헤더에서 열 위치 찾기: 지사명 / 목표금액 / (선택)적용월
  //    적용월 열이 있으면 "그 달부터 적용"되는 월별 목표로 인식한다.
  var nameCol = -1, goalCol = -1, monthCol = -1, headerRow = -1;
  for (var r = 0; r < Math.min(data.length, 5); r++) {
    var nC = -1, gC = -1, mC = -1;
    for (var c = 0; c < data[r].length; c++) {
      var h = norm(data[r][c]);
      if (h.indexOf('지사명') >= 0 || h === '지사') nC = c;
      else if (h.indexOf('목표금액') >= 0 || h.indexOf('목표') >= 0) gC = c;
      if (h.indexOf('적용월') >= 0 || h.indexOf('적용시작월') >= 0 || h.indexOf('기준월') >= 0 || h === '월' || h.indexOf('년월') >= 0) mC = c;
    }
    if (nC >= 0 && gC >= 0) { headerRow = r; nameCol = nC; goalCol = gC; monthCol = mC; break; }
  }
  if (headerRow < 0) { headerRow = 0; nameCol = 0; goalCol = 1; monthCol = -1; }  // A1/B1 기본 가정

  // 2) 지사명 열에서 '헤리티지' 행 찾기
  var found = false;
  for (var r2 = headerRow + 1; r2 < data.length; r2++) {
    var nm = norm(data[r2][nameCol]);
    if (!nm || nm.indexOf(BRANCH_GOAL_NAME) < 0) continue;
    var val = num(data[r2][goalCol]);
    if (val === null || val <= 0) continue;
    out.branch = String(data[r2][nameCol]);
    if (monthCol >= 0) {
      var mk = normalizeMonth(data[r2][monthCol]);
      if (mk) { out.map[mk] = val; found = true; continue; }
    }
    if (out.single === null) { out.single = val; found = true; }
  }
  if (found) {
    out.ok = true;
    out.how = (Object.keys(out.map).length > 0) ? '적용월별 목표' : '단일 목표';
    return out;
  }

  // 3) 헤더 없이 A열 '헤리티지' / B열 금액인 경우
  for (var r3 = 0; r3 < data.length; r3++) {
    if (norm(data[r3][0]).indexOf(BRANCH_GOAL_NAME) >= 0) {
      var v3 = num(data[r3][1]);
      if (v3 !== null && v3 > 0) {
        out.ok = true; out.single = v3; out.how = 'A열/B열';
        out.branch = String(data[r3][0]);
        return out;
      }
    }
  }
  out.reason = "'" + BRANCH_GOAL_NAME + "' 행 또는 목표금액 숫자를 찾지 못함";
  return out;
}

function getBranchGoalData() {
  var last = { ok: false, map: {}, single: null, how: '', reason: '', source: '', branch: '' };
  // 1순위: 별도 지사목표 문서
  try {
    var res1 = _readBranchGoalFromSpreadsheet(SpreadsheetApp.openById(BRANCH_GOAL_SHEET_ID), '지사목표 문서');
    if (res1.ok) return res1;
    last = res1;
  } catch (e) {
    last.reason = '지사목표 문서 열기 실패: ' + String(e);
    last.source = '지사목표 문서';
  }
  // 2순위: 앱 데이터 시트에도 '지사목표' 탭이 있으면 사용
  try {
    var res2 = _readBranchGoalFromSpreadsheet(SpreadsheetApp.openByUrl(SHEET_URL), '앱 데이터 시트');
    if (res2.ok) return res2;
  } catch (e2) { /* 무시하고 1순위 사유 반환 */ }
  return last;
}
// ===== [지사목표] 끝 =====

// [목표달성률] 해당월 FC 목표 로드
// ===== [목표달성률 그룹] FC 그룹 배정 저장/로드 =====
// 그룹: 마스터본부 / 아너스지점 / 지사직할. 월과 무관한 설정이므로 '프로모션설정' 시트를 키로 공유한다.
// 형식: { "마스터본부": ["FC1","FC2"], "아너스지점": [...], "지사직할": [...] }
function getFcGroupConfig() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('프로모션설정');
    if (!sheet) return '{}';
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'fcGroupConfig') {
        return String(data[i][1] || '{}');
      }
    }
    return '{}';
  } catch (e) {
    return '{}';
  }
}

function saveFcGroupConfig(configJson) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('프로모션설정');
    if (!sheet) {
      sheet = ss.insertSheet('프로모션설정');
      sheet.appendRow(['키', '값', '최종수정일']);
    }
    var data = sheet.getDataRange().getValues();
    var ts = new Date();
    var found = false;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'fcGroupConfig') {
        sheet.getRange(i + 1, 2).setValue(configJson);
        sheet.getRange(i + 1, 3).setValue(ts);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow(['fcGroupConfig', configJson, ts]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
// ===== [목표달성률 그룹] 끝 =====

function getFcGoals(month) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('FC목표');
    if (!sheet) return {};
    var data = sheet.getDataRange().getValues();
    var m = normalizeMonth(month);
    var goals = {};
    for (var i = 1; i < data.length; i++) {
      if (normalizeMonth(data[i][0]) === m && data[i][1]) {
        goals[String(data[i][1]).trim()] = Number(data[i][2]) || 0;
      }
    }
    return goals;
  } catch (e) {
    return {};
  }
}

// [목표달성률] FC 목표 저장 (upsert, 시트 자동생성)
function saveFcGoals(month, goalsJson) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('FC목표');
    if (!sheet) {
      sheet = ss.insertSheet('FC목표');
      sheet.appendRow(['기준월', 'FC명', '목표환산P', '최종수정일']);
    }
    var goals = (typeof goalsJson === 'string') ? JSON.parse(goalsJson) : (goalsJson || {});
    var m = normalizeMonth(month);
    var data = sheet.getDataRange().getValues();
    var rowMap = {};
    for (var i = 1; i < data.length; i++) {
      if (normalizeMonth(data[i][0]) === m && data[i][1]) {
        rowMap[String(data[i][1]).trim()] = i + 1;
      }
    }
    var ts = new Date();
    var appends = [];
    for (var name in goals) {
      var val = Number(goals[name]) || 0;
      if (rowMap[name]) {
        sheet.getRange(rowMap[name], 3).setValue(val);
        sheet.getRange(rowMap[name], 4).setValue(ts);
      } else {
        appends.push(["'" + m, name, val, ts]);
      }
    }
    if (appends.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, 4).setValues(appends);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// [특별 프로모션] 프로모션 이름 로드
function getPromoName() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('프로모션설정');
    if (!sheet) return '';
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'promoName') {
        return String(data[i][1] || '');
      }
    }
    return '';
  } catch (e) {
    return '';
  }
}

// [특별 프로모션] 설정 전체(JSON) 로드 - 제목/업적기간/트립명 등
function getPromoConfig() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('프로모션설정');
    if (!sheet) return '{}';
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'promoConfig') {
        return String(data[i][1] || '{}');
      }
    }
    return '{}';
  } catch (e) {
    return '{}';
  }
}

// [특별 프로모션] 설정 전체(JSON) 저장 (upsert)
function savePromoConfig(configJson) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('프로모션설정');
    if (!sheet) {
      sheet = ss.insertSheet('프로모션설정');
      sheet.appendRow(['키', '값', '최종수정일']);
    }
    var data = sheet.getDataRange().getValues();
    var ts = new Date();
    var found = false;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'promoConfig') {
        sheet.getRange(i + 1, 2).setValue(configJson);
        sheet.getRange(i + 1, 3).setValue(ts);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow(['promoConfig', configJson, ts]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// [특별 프로모션] 프로모션 이름 저장 (시트 자동생성, upsert)
function savePromoName(name) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('프로모션설정');
    if (!sheet) {
      sheet = ss.insertSheet('프로모션설정');
      sheet.appendRow(['키', '값', '최종수정일']);
    }
    var data = sheet.getDataRange().getValues();
    var ts = new Date();
    var found = false;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'promoName') {
        sheet.getRange(i + 1, 2).setValue(name);
        sheet.getRange(i + 1, 3).setValue(ts);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow(['promoName', name, ts]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}



function getMonthlyRates() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('월별시책');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0] && row[1]) {
        var monthStr = "";
        if (row[0] instanceof Date) {
          var y = row[0].getFullYear();
          var m = row[0].getMonth() + 1;
          monthStr = y + '-' + (m < 10 ? '0' + m : m);
        } else {
          monthStr = String(row[0]).trim().substring(0, 7);
        }
        result.push({
          month:   monthStr,
          company: String(row[1]).trim(),
          product: String(row[2]).trim(),
          bonus:   row[3],
          bonus13: row[4]
        });
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

function saveUserData(userName, month, jsonData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('DB');
    if (!sheet) {
      sheet = ss.insertSheet('DB');
      sheet.appendRow(['사용자명', '기준월', '데이터(JSON)', '최종수정일']);
    }
    var data = sheet.getDataRange().getValues();
    var targetRow = -1;
    var searchMonth = normalizeMonth(month);
    for (var i = 1; i < data.length; i++) {
      var rowName  = String(data[i][0]).trim();
      var rowMonth = normalizeMonth(data[i][1]);
      if (rowName === userName && rowMonth === searchMonth) {
        targetRow = i + 1;
        break;
      }
    }
    var timestamp = new Date();
    if (targetRow > 0) {
      sheet.getRange(targetRow, 3).setValue(jsonData);
      sheet.getRange(targetRow, 4).setValue(timestamp);
    } else {
      sheet.appendRow([userName, "'" + searchMonth, jsonData, timestamp]);
    }
    // [고객 관리 시트 자동 동기화] 방금 저장한 FC의 탭을 갱신 (탭 없으면 생성)
    // 실패해도 청약 저장 자체는 성공 처리한다.
    try { exportOneFcToSheet(userName, true); } catch (e) {}
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getAllData() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('DB');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var resultList = [];
    for (var i = 1; i < data.length; i++) {
      var name     = String(data[i][0]);
      var rawMonth = data[i][1];
      var json     = data[i][2];
      if (name && rawMonth) {
        resultList.push({
          name:  name,
          month: normalizeMonth(rawMonth),
          json:  json
        });
      }
    }
    return resultList;
  } catch (e) {
    return [];
  }
}

// ====================================================================================
// [4] 게시판 및 공지사항 관련 함수
// ====================================================================================

function getNotices() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('공지사항');
    if (!sheet) {
      sheet = ss.insertSheet('공지사항');
      sheet.appendRow(['날짜', '제목', '내용', '작성자']);
      return [];
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      result.push({
        date:    formatDate(data[i][0]),
        title:   String(data[i][1]),
        content: String(data[i][2]),
        writer:  String(data[i][3])
      });
    }
    return result.reverse();
  } catch (e) { return []; }
}

function getSchedules() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('교육일정');
    if (!sheet) {
      sheet = ss.insertSheet('교육일정');
      sheet.appendRow(['날짜', '시간', '교육명', '강사', '장소']);
      return [];
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      result.push({
        date:       formatDate(data[i][0]),
        time:       formatTime(data[i][1]),
        title:      String(data[i][2]),
        instructor: String(data[i][3]),
        place:      String(data[i][4])
      });
    }
    return result;
  } catch (e) { return []; }
}

function formatTime(val) {
  if (val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    return (h < 10 ? '0' + h : h) + ':' + (m < 10 ? '0' + m : m);
  }
  return String(val);
}

function getBoardData() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('정보게시판');
    if (!sheet) {
      sheet = ss.insertSheet('정보게시판');
      sheet.appendRow(['ID', '날짜', '작성자', '제목', '내용', '이미지데이터']);
      return [];
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var result = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      if (row[0]) {
        result.push({
          id:      row[0],
          date:    formatDate(row[1]),
          writer:  row[2],
          title:   row[3],
          content: row[4],
          image:   row[5] ? row[5].toString() : ""
        });
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

function saveBoardPost(postData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!BOARD_FOLDER_ID || BOARD_FOLDER_ID.includes('붙여넣으세요')) {
      return { success: false, error: "Code.gs 파일 맨 위에 '게시판 폴더 ID'를 설정해주세요." };
    }
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('정보게시판');
    if (!sheet) {
      sheet = ss.insertSheet('정보게시판');
      sheet.appendRow(['ID', '날짜', '작성자', '제목', '내용', '이미지데이터']);
    }
    var timestamp    = new Date();
    var id           = new Date().getTime().toString();
    var finalFilesJson = "";
    if (postData.image && postData.image.length > 5) {
      try {
        finalFilesJson = processAttachments(postData.image);
      } catch (uploadErr) {
        return { success: false, error: "파일 업로드 실패: " + uploadErr.toString() };
      }
    }
    sheet.appendRow([id, timestamp, postData.writer, postData.title, postData.content, finalFilesJson]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateBoardPost(postData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('정보게시판');
    if (!sheet) return { success: false, error: '게시판 시트가 없습니다.' };
    var data = sheet.getDataRange().getDisplayValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(postData.id).trim()) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow === -1) return { success: false, error: '수정할 게시글을 찾을 수 없습니다.' };
    sheet.getRange(targetRow, 4).setValue(postData.title);
    sheet.getRange(targetRow, 5).setValue(postData.content);
    if (postData.image) {
      try {
        var finalFilesJson = processAttachments(postData.image);
        sheet.getRange(targetRow, 6).setValue(finalFilesJson);
      } catch (uploadErr) {
        return { success: false, error: "파일 수정 실패: " + uploadErr.toString() };
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function processAttachments(jsonString) {
  var files          = JSON.parse(jsonString);
  var processedFiles = [];
  var folder         = DriveApp.getFolderById(BOARD_FOLDER_ID);
  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    if (file.data && !file.data.startsWith('data:')) {
      processedFiles.push(file);
      continue;
    }
    if (file.data && file.data.startsWith('data:')) {
      var contentType = file.data.substring(5, file.data.indexOf(';'));
      var base64      = file.data.substring(file.data.indexOf(',') + 1);
      var blob        = Utilities.newBlob(Utilities.base64Decode(base64), contentType, file.name);
      var driveFile   = folder.createFile(blob);
      try {
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        console.log("권한 설정 오류(무시 가능): " + e.toString());
      }
      processedFiles.push({
        type: file.type,
        name: file.name,
        id:   driveFile.getId(),
        data: driveFile.getDownloadUrl()
      });
    }
  }
  return JSON.stringify(processedFiles);
}

function deleteBoardPost(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('정보게시판');
    if (!sheet) return { success: false, error: '시트 없음' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: '게시글을 찾을 수 없습니다.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ====================================================================================
// [5] 기타 유틸리티 함수
// ====================================================================================

function getContactData() {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName('연락처');
    if (!sheet) {
      sheet = ss.insertSheet('연락처');
      sheet.appendRow(['구분', '회사명', '직책', '성명', '연락처']);
      return [];
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var contacts = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[1] && row[3]) {
        contacts.push({
          category: String(row[0]).trim(),
          company:  String(row[1]).trim(),
          role:     String(row[2]).trim(),
          name:     String(row[3]).trim(),
          phone:    String(row[4]).trim()
        });
      }
    }
    return contacts;
  } catch (e) {
    return [];
  }
}

function getDriveLinks() {
  return {
    life: LINK_LIFE_INS
  };
}

function getLifeInsuranceFiles() {
  try {
    var folder = DriveApp.getFolderById(LIFE_INS_FOLDER_ID);
    var files  = folder.getFiles();
    var fileList = [];
    while (files.hasNext()) {
      var file     = files.next();
      var mimeType = file.getMimeType();
      if (mimeType.includes('image') || mimeType.includes('pdf')) {
        fileList.push({
          id:          file.getId(),
          name:        file.getName(),
          type:        mimeType.includes('pdf') ? 'pdf' : 'image',
          mimeType:    mimeType,
          lastUpdated: formatDate(file.getLastUpdated())
        });
      }
    }
    return fileList.sort(function(a, b) { return a.name.localeCompare(b.name); });
  } catch (e) {
    return { error: true, message: e.toString() };
  }
}

function getFileBase64Data(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    throw new Error("파일을 불러올 수 없습니다.");
  }
}

function getFinancialAppHtml() {
  return HtmlService.createHtmlOutputFromFile('FinancialPlan').getContent();
}

// ── 헬퍼: 나이 계산 ──────────────────────────────────
function calculateAge(birthDate) {
  if (!birthDate || birthDate.length !== 8) return 0;
  const year  = birthDate.substring(0, 4);
  const month = birthDate.substring(4, 6);
  const day   = birthDate.substring(6, 8);
  const today = new Date();
  const birth = new Date(year, month - 1, day);
  let age = today.getFullYear() - birth.getFullYear();
  const monthDiff = today.getMonth() - birth.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birth.getDate())) {
    age--;
  }
  return age;
}

// ── 헬퍼: 생년월일 포맷 변환 ────────────────────────
function formatBirthDate(birthDate) {
  const year  = birthDate.substring(0, 4);
  const month = birthDate.substring(4, 6);
  const day   = birthDate.substring(6, 8);
  return year + '-' + month + '-' + day;
}

// ── 메인 진입점 ──────────────────────────────────────
function processFinancialPlanning(formData) {
  try {
    const age                = calculateAge(formData.birthDate);
    const formattedBirthDate = formatBirthDate(formData.birthDate);
    const cleanSalary  = Number(String(formData.salary).replace(/,/g, ''));
    const cleanExpense = Number(String(formData.expense).replace(/,/g, ''));

    const userData = {
      name:            formData.name,
      birthDate:       formattedBirthDate,
      age:             age,
      salary:          cleanSalary,
      marriage:        formData.marriage,
      children:        formData.children,
      housing:         formData.housing,
      goal:            formData.goal,
      expense:         cleanExpense,
      // ✅ 직접 배분값 전달 (최적화 버튼 사용 시)
      customLiving:    formData.customLiving,
      customSaving:    formData.customSaving,
      customEmergency: formData.customEmergency,
      customInsurance: formData.customInsurance
    };

    let reportData;
    if (userData.goal === '대출상환' || userData.goal === '노후준비') {
      reportData = generateAiBasedData(userData);
    } else {
      reportData = generateRuleBasedData(userData);
    }

    saveToFinancialSheet(userData, reportData);
    return { success: true, data: reportData, userInfo: userData };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ── 규칙 기반 계산 (직접배분 분기 적용) ─────────────
function generateRuleBasedData(userData) {
  const salary    = Number(userData.salary);
  const expense   = Number(userData.expense);
  const available = salary - expense;

  // 직접 입력값이 있으면 그 값을, 없으면 기본 비율 적용
  const livingCost    = userData.customLiving    !== undefined ? Number(userData.customLiving)    : Math.round(available * 0.3);
  const savingCost    = userData.customSaving    !== undefined ? Number(userData.customSaving)    : Math.round(available * 0.5);
  const emergencyCost = userData.customEmergency !== undefined ? Number(userData.customEmergency) : Math.round(available * 0.05);
  const insuranceCost = userData.customInsurance !== undefined ? Number(userData.customInsurance) : Math.round(available * 0.15);

  // 저축 세부 항목은 확정된 savingCost 기준 비율 적용
  const bankSavings     = Math.round(savingCost * 0.25);
  const fundSavings     = Math.round(savingCost * 0.25);
  const midTermSavings  = Math.round(savingCost * 0.30);
  const longTermSavings = Math.round(savingCost * 0.20);

  const analysisText = `
${userData.name} 고객님은 현재 ${userData.age}세 ${userData.marriage}으로, 자녀 ${userData.children}과 함께 ${userData.housing}에 거주하고 계십니다. 
월 세후 급여 ${salary.toLocaleString()}원 중 고정지출을 제외한 가용자금은 <strong>${available.toLocaleString()}원</strong>입니다. 
현재 가장 중요한 목표인 <strong>'${userData.goal}'</strong> 달성을 위해, 소득의 일정 비율 이상을 저축/투자에 집중하는 '선저축 후지출' 시스템 구축이 시급합니다.
`;

  const riskText = `
인생의 5대 자금(결혼, 주택, 교육, 노후, 의료비) 중 가장 예측 불가능한 의료비 리스크를 헷지해야 합니다. 
적정보험료 <strong>${insuranceCost.toLocaleString()}원</strong> 한도 내에서 기존 보장을 점검하고, 가장의 유고 시 가족의 생활을 지킬 수 있는 안전장치를 마련하시기 바랍니다.
<br><span style="color:#d97706; font-weight:bold;">※ 자세한 상담은 전문가와 함께하세요.</span>
`;

  return {
    analysis: analysisText.trim(),
    risk:     riskText.trim(),
    allocation: {
      living:    livingCost,
      saving:    savingCost,
      emergency: emergencyCost,
      insurance: insuranceCost,
      detail: {
        bank:     bankSavings,
        fund:     fundSavings,
        midTerm:  midTermSavings,
        longTerm: longTermSavings
      }
    }
  };
}

// ── AI 기반 계산 (대출상환·노후준비, 프롬프트 개선) ──
function generateAiBasedData(userData) {
  try {
    const salary    = Number(userData.salary);
    const expense   = Number(userData.expense);
    const available = salary - expense;

    // 직접 입력값 우선 적용
    const livingCost    = userData.customLiving    !== undefined ? Number(userData.customLiving)    : Math.round(available * 0.3);
    const savingCost    = userData.customSaving    !== undefined ? Number(userData.customSaving)    : Math.round(available * 0.5);
    const emergencyCost = userData.customEmergency !== undefined ? Number(userData.customEmergency) : Math.round(available * 0.05);
    const insuranceCost = userData.customInsurance !== undefined ? Number(userData.customInsurance) : Math.round(available * 0.15);

    const bankSavings     = Math.round(savingCost * 0.25);
    const fundSavings     = Math.round(savingCost * 0.25);
    const midTermSavings  = Math.round(savingCost * 0.30);
    const longTermSavings = Math.round(savingCost * 0.20);

    const prompt = `
당신은 최고의 재무 컨설턴트입니다. 아래 고객 정보를 바탕으로 '${userData.goal}' 목표에 특화된 재무설계 JSON 데이터를 만들어주세요.

[고객 정보]
이름: ${userData.name}, 나이: ${userData.age}, 결혼: ${userData.marriage}, 자녀: ${userData.children}
월소득: ${salary}, 고정지출: ${expense}, 가용자금: ${available}
목표: ${userData.goal}

[필수 지시사항]
- "allocation" 항목 내의 모든 숫자는 아래 제공된 값을 절대 변경하지 말고 그대로 출력하세요.
- 응답은 오직 JSON 문자열만 보내주세요. 마크다운(\`\`\`) 금지.

[필수 JSON 포맷]
{
  "analysis": "HTML 태그가 포함된 분석 내용. 가용자금과 목표를 언급하며 전문적으로 작성. 주요 숫자는 <strong>태그 사용.",
  "risk": "HTML 태그가 포함된 리스크 관리 조언. 마지막에 <br><span style='color:#d97706; font-weight:bold;'>※ 자세한 상담은 전문가와 함께하세요.</span> 포함.",
  "allocation": {
    "living": ${livingCost},
    "saving": ${savingCost},
    "emergency": ${emergencyCost},
    "insurance": ${insuranceCost},
    "detail": {
      "bank": ${bankSavings},
      "fund": ${fundSavings},
      "midTerm": ${midTermSavings},
      "longTerm": ${longTermSavings}
    }
  }
}
`;

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;
    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7 }
    };
    const options = {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response     = UrlFetchApp.fetch(url, options);
    const jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse.candidates && jsonResponse.candidates[0].content) {
      let rawText = jsonResponse.candidates[0].content.parts[0].text;
      rawText = rawText.replace(/```json/g, '').replace(/```/g, '').trim();
      return JSON.parse(rawText);
    } else {
      throw new Error('AI 응답 없음');
    }
  } catch (error) {
    Logger.log('AI Error: ' + error.toString());
    return generateRuleBasedData(userData);
  }
}

// ── 시트 저장 ─────────────────────────────────────────
function saveToFinancialSheet(userData, reportData) {
  try {
    const ss = SpreadsheetApp.openById(FINANCIAL_SHEET_ID);
    let sheet = ss.getSheetByName(FINANCIAL_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(FINANCIAL_SHEET_NAME);
      const headers = ['타임스탬프', '이름', '생년월일', '나이', '세후급여', '결혼유무', '자녀유무', '주거형태', '재무목표', '고정지출', '가용자금'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a5568').setFontColor('#ffffff');
    }
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const available = Number(userData.salary) - Number(userData.expense);
    const rowData   = [timestamp, userData.name, userData.birthDate, userData.age, userData.salary, userData.marriage, userData.children, userData.housing, userData.goal, userData.expense, available];
    sheet.appendRow(rowData);
    return { success: true };
  } catch (error) {
    throw new Error('데이터 저장 실패: ' + error.message);
  }
}

// ====================================================================================
// [6] 목표 관리 (Goal) 관련 함수
// [수정] 시트 컬럼 순서를 [기준월(0), 성명(1), 목표금액(2), 수정일(3)]로 통일
// ====================================================================================

// [목표달성률] saveTargetToSheet 제거 - saveFcGoals로 대체


// [목표달성률] getTargetsByMonth 제거 - getFcGoals로 대체


// ====================================================================================
// [8] 전사 통합 랭킹
// ====================================================================================

const EXTERNAL_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1LejP4uKQoYBmTkjgYObXBaPzE29FJdMgONmcyO-3YA4/edit';

// [신채널 순위] 지사별 메타 정보: 지사목표 + 재적 인원
//  - goals: '지사목표' 시트(외부 문서 우선) 전 지사. '적용월' 열 있으면 조회월 이하 최신 행(carry-forward).
//  - members: 본점(헤리티지) + 외부 '사용자 정보'의 지사별 재적 수(직책·지점 있고 조회월 재직).
function getGlobalRankingMeta(targetMonth) {
  var meta = { goals: {}, members: {} };
  var monthStr = String(targetMonth || '').trim().substring(0, 7);

  // ----- 1. 지사목표 (전 지사) -----
  function readGoals(ss) {
    var sheet = ss.getSheetByName('지사목표');
    if (!sheet) return false;
    var rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return false;
    // 헤더 탐지
    var hIdx = -1, cName = 0, cGoal = 1, cMonth = -1;
    for (var i = 0; i < Math.min(rows.length, 5); i++) {
      for (var j = 0; j < rows[i].length; j++) {
        var v = String(rows[i][j]).replace(/\s/g, '');
        if (v.indexOf('지사명') >= 0) { hIdx = i; cName = j; }
        if (v.indexOf('목표') >= 0 && hIdx === i) cGoal = j;
        if (v.indexOf('적용월') >= 0 && hIdx === i) cMonth = j;
      }
      if (hIdx >= 0) break;
    }
    if (hIdx < 0) { hIdx = 0; }
    var found = false;
    var best = {};   // 지사명 → { m: 적용월, goal }
    for (var r = hIdx + 1; r < rows.length; r++) {
      var bn = String(rows[r][cName] || '').trim();
      if (!bn) continue;
      var g = parseFloat(String(rows[r][cGoal]).replace(/[^0-9.\-]/g, '')) || 0;
      if (g <= 0) continue;
      var am = '';
      if (cMonth >= 0 && rows[r][cMonth] !== '' && rows[r][cMonth] !== null) {
        var mv = rows[r][cMonth];
        if (Object.prototype.toString.call(mv) === '[object Date]') {
          am = mv.getFullYear() + '-' + ('0' + (mv.getMonth() + 1)).slice(-2);
        } else {
          am = String(mv).trim().replace(/[./]/g, '-').substring(0, 7);
        }
      }
      // carry-forward: 적용월 없으면 기본, 있으면 조회월 이하 최신
      var cur = best[bn];
      if (!am) {
        if (!cur) best[bn] = { m: '', goal: g };
      } else if (am <= monthStr) {
        if (!cur || cur.m === '' || am >= cur.m) best[bn] = { m: am, goal: g };
      }
      found = true;
    }
    for (var k in best) meta.goals[k] = best[k].goal;
    return found;
  }
  try {
    var ok = false;
    try { ok = readGoals(SpreadsheetApp.openById(BRANCH_GOAL_SHEET_ID)); } catch (e1) { ok = false; }
    if (!ok) { try { readGoals(SpreadsheetApp.openByUrl(SHEET_URL)); } catch (e2) {} }
  } catch (e) {}

  // ----- 2. 재적 인원 -----
  function countMembers(ss, defaultBranch) {
    var sheet = ss.getSheetByName('사용자 정보');
    if (!sheet) return;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var name = String(rows[i][0] || '').trim();
      var role = String(rows[i][1] || '').trim();
      var br   = String(rows[i][2] || '').trim();
      var resign = rows[i][4];
      if (!name || !role || !br) continue;
      // 조회월 재직 판정
      if (resign !== '' && resign !== null && resign !== undefined) {
        var rs = '';
        if (Object.prototype.toString.call(resign) === '[object Date]') {
          rs = resign.getFullYear() + '-' + ('0' + (resign.getMonth() + 1)).slice(-2);
        } else {
          rs = String(resign).trim().substring(0, 7);
        }
        if (rs && rs < monthStr) continue;
      }
      var bKey = defaultBranch || br;
      meta.members[bKey] = (meta.members[bKey] || 0) + 1;
    }
  }
  // (1) 본점: 헤리티지 재적
  try { countMembers(SpreadsheetApp.openByUrl(SHEET_URL), '헤리티지'); } catch (e3) {}
  // (2) 타 지사: 통합 문서(BRANCH_GOAL_SHEET_ID)의 '사용자 정보' 탭에서 지사명(C열)별 재적
  //     이 문서에는 더드림·일산·춘천·부천·송파 등 전 지사 인원이 들어 있다. 헤리티지는 (1)에서 집계하므로 여기선 제외.
  try {
    var ssMeta = SpreadsheetApp.openById(BRANCH_GOAL_SHEET_ID);
    var mSheet = ssMeta.getSheetByName('사용자 정보');
    if (mSheet) {
      var mRows = mSheet.getDataRange().getValues();
      for (var mi = 1; mi < mRows.length; mi++) {
        var mName = String(mRows[mi][0] || '').trim();
        var mRole = String(mRows[mi][1] || '').trim();
        var mBr   = String(mRows[mi][2] || '').trim();
        var mResign = mRows[mi][4];
        if (!mName || !mRole || !mBr) continue;
        if (mBr.indexOf('헤리티지') >= 0) continue;   // 헤리티지는 본점에서 집계
        if (mResign !== '' && mResign !== null && mResign !== undefined) {
          var mrs = '';
          if (Object.prototype.toString.call(mResign) === '[object Date]') {
            mrs = mResign.getFullYear() + '-' + ('0' + (mResign.getMonth() + 1)).slice(-2);
          } else {
            mrs = String(mResign).trim().substring(0, 7);
          }
          if (mrs && mrs < monthStr) continue;   // 조회월 이전 퇴사자 제외
        }
        meta.members[mBr] = (meta.members[mBr] || 0) + 1;
      }
    }
  } catch (e4) {}

  return meta;
}

function getGlobalRankingData(targetMonth) {
  var allRows      = [];
  var searchString = String(targetMonth).trim();

  // 1. 본점 데이터
  try {
    var ssMain     = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheetMain  = ssMain.getSheetByName('DB');
    var myBranchName = "헤리티지";
    if (sheetMain) {
      var data = sheetMain.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row      = data[i];
        var isMatch  = false;
        var cellDate = row[1];
        if (cellDate) {
          var dateStr = "";
          if (cellDate instanceof Date) {
            var y = cellDate.getFullYear();
            var m = cellDate.getMonth() + 1;
            dateStr = y + '-' + (m < 10 ? '0' + m : m);
          } else {
            dateStr = String(cellDate).trim().substring(0, 7);
          }
          if (dateStr === searchString) isMatch = true;
        }
        if (isMatch) {
          allRows.push({
            branchName: myBranchName,
            userName:   String(row[0]).trim(),
            jsonData:   row[2]
          });
        }
      }
    }
  } catch (e) {
    console.error("본점 데이터 로드 중 오류: " + e.toString());
  }

  // 2. 타 지사 데이터
  try {
    if (EXTERNAL_SHEET_URL && !EXTERNAL_SHEET_URL.includes("https://docs.google.com/spreadsheets/d/1LejP4uKQoYBmTkjgYObXBaPzE29FJdMgONmcyO-3YA4/edit?gid=0#gid=0")) {
      var ssExt     = SpreadsheetApp.openByUrl(EXTERNAL_SHEET_URL);
      var sheetUser = ssExt.getSheetByName('사용자 정보');
      var userBranchMap = {};
      if (sheetUser) {
        var userData = sheetUser.getDataRange().getValues();
        for (var u = 1; u < userData.length; u++) {
          var uName   = String(userData[u][0]).trim();
          var uBranch = String(userData[u][2]).trim();
          if (uName) userBranchMap[uName] = uBranch || "기타 지사";
        }
      }
      var sheetDB = ssExt.getSheetByName('DB');
      if (sheetDB) {
        var dbData = sheetDB.getDataRange().getValues();
        for (var r = 1; r < dbData.length; r++) {
          var extRow   = dbData[r];
          var extMatch = false;
          var extCell  = extRow[1];
          if (extCell) {
            var extDateStr = "";
            if (extCell instanceof Date) {
              var ey = extCell.getFullYear();
              var em = extCell.getMonth() + 1;
              extDateStr = ey + '-' + (em < 10 ? '0' + em : em);
            } else {
              extDateStr = String(extCell).trim().substring(0, 7);
            }
            if (extDateStr === searchString) extMatch = true;
          }
          if (extMatch) {
            var extUserName   = String(extRow[0]).trim();
            var extBranchName = userBranchMap[extUserName] || "타 지사(미분류)";
            allRows.push({
              branchName: extBranchName,
              userName:   extUserName,
              jsonData:   extRow[2]
            });
          }
        }
      }
    }
  } catch (e) {
    console.error("타 지사 시트 로드 중 오류: " + e.toString());
  }

  return allRows;
}

function getSonboWeekConfig() {
  try {
    const ss    = SpreadsheetApp.openById('1Hd5bMxVTycTRd2SOr8MrhwOFF3dbCAHyNDxiWUgXvsk');
    const sheet = ss.getSheetByName('손보주차');
    if (!sheet) {
      Logger.log('손보주차 시트를 찾을 수 없습니다');
      return [];
    }
    const data   = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      let monthStr = '';
      if (row[0] instanceof Date) {
        const y = row[0].getFullYear();
        const m = String(row[0].getMonth() + 1).padStart(2, '0');
        monthStr = y + '-' + m;
      } else {
        monthStr = String(row[0]).trim().substring(0, 7);
      }
      result.push({
        month: monthStr,
        w1s: Number(row[1]) || 1,
        w1e: Number(row[2]) || 7,
        w2s: Number(row[3]) || 8,
        w2e: Number(row[4]) || 14,
        w3s: Number(row[5]) || 15,
        w3e: Number(row[6]) || 21,
        w4s: Number(row[7]) || 22,
        w4e: Number(row[8]) || 31,
      });
    }
    Logger.log('손보주차 결과: ' + JSON.stringify(result));
    return result;
  } catch(e) {
    Logger.log('getSonboWeekConfig 오류: ' + e.toString());
    return [];
  }
}

// [삭제됨] getStrategicAdvice (AI전략) - 상품전략(PDF)으로 대체


function getPhotoGallery() {
  try {
    const folder    = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    const files     = folder.getFiles();
    const photoList = [];
    while (files.hasNext()) {
      const file     = files.next();
      const mimeType = file.getMimeType();
      if (mimeType.startsWith('image/')) {
        photoList.push({
          id:      file.getId(),
          name:    file.getName(),
          date:    formatDate(file.getLastUpdated()),
          url:     `https://drive.google.com/thumbnail?sz=w400&id=${file.getId()}`,
          fullUrl: `https://drive.google.com/file/d/${file.getId()}/view`
        });
      }
    }
    photoList.sort((a, b) => new Date(b.date) - new Date(a.date));
    return photoList;
  } catch(e) {
    return { error: true, message: e.toString() };
  }
}

function getSilsonRates() {
  try {
    const ss    = SpreadsheetApp.openById('1Hd5bMxVTycTRd2SOr8MrhwOFF3dbCAHyNDxiWUgXvsk');
    const sheet = ss.getSheetByName('실손수수료');
    if (!sheet) return {};
    const rows   = sheet.getDataRange().getValues();
    const result = {};
    rows.forEach((row, i) => {
      if (i === 0) return;
      const company     = String(row[0]).trim();
      const nextRate    = parseFloat(row[1]) || 0;
      const bonus13Rate = parseFloat(row[2]) || 0;
      if (company) {
        result[company] = { nextMonth: nextRate, bonus13: bonus13Rate };
      }
    });
    return result;
  } catch(e) {
    Logger.log('getSilsonRates 오류: ' + e);
    return {};
  }
}

// [개편-신규] 손보장기 회사별 시책률 로드 (월별 지원)
//  - 구글시트 '손보장기시책' 탭: A열 회사명 / B열 1·2주차 시책률(%) / C열 3·4주차 시책률(%) / D열 적용월(YYYY-MM, 선택)
//  - 시책률 셀은 퍼센트(%) 서식. getValues는 소수값(600% → 6.0) 반환.
//  - [월별] 같은 회사를 여러 행에 적용월과 함께 넣으면, 계약월 이하의 가장 최근 적용월 행이 적용됨.
//    적용월이 빈 행은 기본값(어느 적용월에도 해당 안 될 때 사용).
//  - '기타' 행을 두면 목록에 없는 회사의 기본값으로 사용됨. 탭 없으면 자동 생성.
function getSonboLongRates() {
  try {
    const ss    = SpreadsheetApp.openById('1Hd5bMxVTycTRd2SOr8MrhwOFF3dbCAHyNDxiWUgXvsk');
    let   sheet = ss.getSheetByName('손보장기시책');
    if (!sheet) {
      sheet = ss.insertSheet('손보장기시책');
      sheet.appendRow(['회사명', '1·2주차 시책률(%)', '3·4주차 시책률(%)', '적용월']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#0d9488').setFontColor('#ffffff');
      sheet.getRange(2, 2, 100, 2).setNumberFormat('0%');
      sheet.appendRow(['기타', 6.0, 4.5, '']);
      return { '기타': [{ month: '', rate12: 6.0, rate34: 4.5 }] };
    }
    const rows   = sheet.getDataRange().getValues();
    const result = {};
    for (let i = 1; i < rows.length; i++) {
      const company = String(rows[i][0]).trim();
      if (!company) continue;
      const rate12 = parseFloat(rows[i][1]) || 0;
      const rate34 = parseFloat(rows[i][2]) || 0;
      // D열 적용월: 날짜 셀/텍스트 모두 지원 → 'YYYY-MM'으로 정규화
      let month = '';
      if (rows[i].length > 3 && rows[i][3] !== '' && rows[i][3] !== null && rows[i][3] !== undefined) {
        const mv = rows[i][3];
        if (Object.prototype.toString.call(mv) === '[object Date]') {
          month = mv.getFullYear() + '-' + ('0' + (mv.getMonth() + 1)).slice(-2);
        } else {
          month = String(mv).trim().replace(/[./]/g, '-').substring(0, 7);
        }
      }
      if (!result[company]) result[company] = [];
      result[company].push({ month: month, rate12: rate12, rate34: rate34 });
    }
    return result;
  } catch (e) {
    Logger.log('getSonboLongRates 오류: ' + e);
    return {};
  }
}

// ====================================================================================
// [9] 리크루팅 관련 함수
// ====================================================================================

var SPREADSHEET_ID = "1Hd5bMxVTycTRd2SOr8MrhwOFF3dbCAHyNDxiWUgXvsk";

function getRecruitingData() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("리크루팅DB");
    if (!sheet) return [];
    var data = sheet.getRange("A1").getValue();
    return data ? JSON.parse(data) : [];
  } catch (e) {
    return [];
  }
}

function saveRecruitingData(jsonStr) {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet   = ss.getSheetByName("리크루팅DB") || ss.insertSheet("리크루팅DB");
    var folder  = DriveApp.getFolderById(BOARD_FOLDER_ID);
    var recruits = JSON.parse(jsonStr);

    recruits = recruits.map(function(recruit) {
      if (recruit.cisFile && recruit.cisFile.data && recruit.cisFile.data.startsWith('data:')) {
        try { recruit.cisFile = uploadRecruitFile(recruit.cisFile, folder); }
        catch(e) { Logger.log('cisFile 오류: ' + e); }
      }
      if (recruit.interviewFile && recruit.interviewFile.data && recruit.interviewFile.data.startsWith('data:')) {
        try { recruit.interviewFile = uploadRecruitFile(recruit.interviewFile, folder); }
        catch(e) { Logger.log('interviewFile 오류: ' + e); }
      }
      if (recruit.commissionFiles && Array.isArray(recruit.commissionFiles)) {
        recruit.commissionFiles = recruit.commissionFiles.map(function(f) {
          if (f && f.data && f.data.startsWith('data:')) {
            try { return uploadRecruitFile(f, folder); } catch(e) { return f; }
          }
          return f;
        });
      }
      return recruit;
    });

    sheet.getRange("A1").setValue(JSON.stringify(recruits));
    return { success: true };
  } catch (e) {
    console.error("저장 에러: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function uploadRecruitFile(fileObj, folder) {
  var dataStr     = fileObj.data;
  var contentType = dataStr.substring(5, dataStr.indexOf(';'));
  var base64      = dataStr.substring(dataStr.indexOf(',') + 1);
  var blob        = Utilities.newBlob(Utilities.base64Decode(base64), contentType, fileObj.name || 'file');
  var driveFile   = folder.createFile(blob);
  try {
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  return {
    name:    fileObj.name || 'file',
    id:      driveFile.getId(),
    viewUrl: 'https://drive.google.com/file/d/' + driveFile.getId() + '/view'
  };
}

// ==========================================
// [신규] 조직도 데이터 가져오기 (하위 부서 지원)
// ==========================================
function getOrgChartData() {
  // 본인의 스프레드시트 ID로 변경하세요
  var SPREADSHEET_ID = '1Hd5bMxVTycTRd2SOr8MrhwOFF3dbCAHyNDxiWUgXvsk'; 
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("조직도");
  
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var result = [];
  
  // 1행은 헤더이므로 i = 1 부터 시작
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][3]) { // 상위소속과 이름이 있는 경우만
      result.push({
        parentDept: data[i][0].toString().trim(),
        subDept: (data[i][1] || '').toString().trim(), // 빈칸 허용
        role: (data[i][2] || '').toString().trim(),
        name: data[i][3].toString().trim()
      });
    }
  }
  return result;
}


// ============================================================
// FC 신규등록 (코인스 프로와 동일 문서/탭/컬럼에 저장)
// 저장 문서: 1LejP4uKQoYBmTkjgYObXBaPzE29FJdMgONmcyO-3YA4 (코인스 프로와 공유)
// 탭: FC신규등록 | 컬럼: id | 지사 | 성명 | 등록예정월 | 등록회차 | 경력무경력 | 최종등록여부 | 등록시각
// ※ 헤리티지 프로 실행 계정이 위 문서에 "편집" 권한이 있어야 저장됩니다.
// ============================================================
var FC_REG_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1LejP4uKQoYBmTkjgYObXBaPzE29FJdMgONmcyO-3YA4/edit';
var FC_REG_TAB = 'FC신규등록';
var FC_REG_HEADERS = ['id', '지사', '성명', '등록예정월', '등록회차', '경력무경력', '최종등록여부', '등록시각'];

// 탭 핸들 확보 (없으면 생성 + 헤더 세팅)
function _getFcRegSheet_() {
  var ss = SpreadsheetApp.openByUrl(FC_REG_SHEET_URL);
  var sh = ss.getSheetByName(FC_REG_TAB);
  if (!sh) {
    sh = ss.insertSheet(FC_REG_TAB);
    sh.getRange(1, 1, 1, FC_REG_HEADERS.length).setValues([FC_REG_HEADERS]);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, FC_REG_HEADERS.length).setValues([FC_REG_HEADERS]);
  }
  return sh;
}

// FC 신규등록 → 위 컬럼 순서로 appendRow, 최종등록여부 'Y' 기본, 생성한 id 반환
function fcRegister(branch, name, ymonth, roundStr, career) {
  var sh = _getFcRegSheet_();
  var id = String(new Date().getTime()) + '_' + Math.floor(1000 + Math.random() * 9000);
  var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([
    id,
    String(branch || ''),
    String(name || ''),
    String(ymonth || ''),
    String(roundStr || ''),
    String(career || ''),
    'N',
    now
  ]);
  return id;
}

// 전체 목록 반환 (헤더 제외) → [{id, branch, name, ymonth, round, career, fin}]
function getFcRegistrations() {
  var sh = _getFcRegSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, FC_REG_HEADERS.length).getValues();
  var result = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    result.push({
      id: String(r[0]),
      branch: String(r[1] || ''),
      name: String(r[2] || ''),
      ymonth: String(r[3] || ''),
      round: String(r[4] || ''),
      career: String(r[5] || ''),
      fin: String(r[6] || '')
    });
  }
  return result;
}

// id로 행 찾아 최종등록여부(7번째 열)를 'Y'/'N'으로 변경
function updateFcFinal(id, yn) {
  var sh = _getFcRegSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.getRange(i + 2, 7).setValue((yn === 'Y') ? 'Y' : 'N');
      return true;
    }
  }
  return false;
}

// id로 행 찾아 삭제
function deleteFcReg(id) {
  var sh = _getFcRegSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

// id로 행 찾아 성명/등록예정월/등록회차/경력무경력/최종등록여부 일괄 수정
function updateFcReg(id, name, ymonth, roundStr, career, fin) {
  var sh = _getFcRegSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      var row = i + 2;
      // 성명(3) 등록예정월(4) 등록회차(5) 경력무경력(6) 최종등록여부(7)
      sh.getRange(row, 3).setValue(String(name || ''));
      sh.getRange(row, 4).setValue(String(ymonth || ''));
      sh.getRange(row, 5).setValue(String(roundStr || ''));
      sh.getRange(row, 6).setValue(String(career || ''));
      sh.getRange(row, 7).setValue((fin === 'Y') ? 'Y' : 'N');
      return true;
    }
  }
  return false;
}


// ============================================================
// 계약 1년 도래 2주 전 알림 메일
// 매일 아침 트리거로 실행 -> 오늘 기준 정확히 14일 뒤 청약 1주년이 되는 계약을
// 담당 FC에게(사용자 정보 탭 이메일), 관리자(지사장/총무)에게 전체 목록을 발송
// ============================================================

// 사용자명 -> 이메일 매핑
function _getEmailMap_() {
  var users = getUserInfo();
  var map = {};
  for (var i = 0; i < users.length; i++) {
    if (users[i].name) map[users[i].name] = { email: users[i].email || '', role: users[i].role || '' };
  }
  return map;
}

// 오늘 기준 daysAhead(기본 14) 뒤 1주년이 도래하는 계약 수집
// startAhead일 뒤부터 span일 동안(기본: 14일 뒤 ~ 20일 뒤, 총 7일)
// 그 기간에 청약 1주년이 도래하는 계약을 모두 수집
function _collectAnniversaryContracts_(startAhead, rangeDays) {
  var start = (startAhead === undefined || startAhead === null) ? 14 : startAhead;
  var span = (rangeDays === undefined || rangeDays === null) ? 7 : rangeDays;
  var tz = 'Asia/Seoul';
  var today = new Date();
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  var windowStart = new Date(today.getTime() + start * 24 * 60 * 60 * 1000);
  var windowEnd = new Date(today.getTime() + (start + span - 1) * 24 * 60 * 60 * 1000);

  var rows = getAllData();
  var results = [];
  for (var i = 0; i < rows.length; i++) {
    var fc = rows[i].name;
    var json = rows[i].json;
    if (!json) continue;
    var items;
    try { items = JSON.parse(json); } catch (e) { continue; }
    if (!items || !items.length) continue;
    for (var k = 0; k < items.length; k++) {
      var it = items[k];
      if (!it || it.type || !it.productKey || it.isWithdrawn || it.isSharedView) continue;
      var dateStr = it.date;
      if (!dateStr || String(dateStr).length < 10) continue;
      var contractYear = parseInt(String(dateStr).substring(0, 4), 10);
      var md = String(dateStr).substring(5, 10);

      var hit = false, annivStr = '';
      for (var y = windowStart.getFullYear(); y <= windowEnd.getFullYear(); y++) {
        var parts = md.split('-');
        var annivDate = new Date(y, parseInt(parts[0], 10) - 1, parseInt(parts[1], 10));
        annivDate = new Date(annivDate.getFullYear(), annivDate.getMonth(), annivDate.getDate());
        if (annivDate >= windowStart && annivDate <= windowEnd && contractYear < y) {
          hit = true;
          annivStr = Utilities.formatDate(annivDate, tz, 'yyyy-MM-dd');
          break;
        }
      }
      if (!hit) continue;

      results.push({
        fc: fc,
        contractor: it.contractor || '(고객명 없음)',
        date: dateStr,
        anniv: annivStr,
        company: it.company || '',
        productName: it.productName || '',
        premium: it.premium || 0
      });
    }
  }
  results.sort(function(a, b){ return String(a.anniv).localeCompare(String(b.anniv)); });
  return results;
}

function sendAnniversaryReminders() {
  var list = _collectAnniversaryContracts_(14, 7);  // 2주 뒤부터 1주일 안에 1주년 도래
  if (!list.length) return '발송 대상 없음';
  var emailMap = _getEmailMap_();
  var byFc = {};
  for (var i = 0; i < list.length; i++) {
    var fc = list[i].fc;
    if (!byFc[fc]) byFc[fc] = [];
    byFc[fc].push(list[i]);
  }
  var sentCount = 0;
  // FC 본인에게만 담당 고객 안내 메일 발송 (1인 1통)
  for (var fcName in byFc) {
    var info = emailMap[fcName];
    if (!info || !info.email) continue;
    var rws = byFc[fcName];
    var titledName = fcName + _roleSuffix_(info.role);
    try {
      MailApp.sendEmail({
        to: info.email,
        subject: '[헤리티지] 계약 1주년 안내 (' + rws.length + '건)',
        htmlBody: _buildFcMailBody_(titledName, rws)
      });
      sentCount++;
    } catch (e) {}
  }
  return '발송 완료: ' + sentCount + '통 (대상 계약 ' + list.length + '건)';
}

// 이름 뒤에 붙일 직급 문자열 (예: ' 지사장님', ' CM님', ' FC님')
function _roleSuffix_(role) {
  var r = String(role || '').trim();
  if (!r) return '님';
  return ' ' + r + '님';
}

function _fmtWon_(n) { n = Number(n) || 0; return n.toLocaleString ? n.toLocaleString() : String(n); }

function _buildFcMailBody_(titledName, rows) {
  var html = '<div style="font-family:Malgun Gothic,sans-serif;color:#222;">';
  html += '<h2 style="color:#0d9488;">' + titledName + ', 곧 1주년이 되는 고객이 있습니다.</h2>';
  html += '<br>';
  html += '<p>아래 고객들은 곧 계약 1주년이 도래합니다. 고객님께 전화드려 CS미팅을 잡아서 계약 리뷰 및 업셀링을 진행해주세요.</p>';
  html += '<br>';
  // [모바일 대응] 고객명·청약일·1주년·보험료는 줄바뀜 없이(nowrap) 작은 폰트로, 보험사/상품만 줄바뀜 허용
  var thBase = 'border:1px solid #ccc;padding:4px 5px;font-size:11px;white-space:nowrap;';
  var tdBase = 'border:1px solid #ccc;padding:4px 5px;font-size:11px;white-space:nowrap;';
  html += '<table style="border-collapse:collapse;width:100%;table-layout:auto;">';
  html += '<tr style="background:#f0fdfa;">'
       + '<th style="' + thBase + '">고객명</th>'
       + '<th style="' + thBase + '">청약일</th>'
       + '<th style="' + thBase + 'white-space:normal;">보험사/상품</th>'
       + '<th style="' + thBase + '">보험료</th></tr>';
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    html += '<tr>'
      + '<td style="' + tdBase + 'font-weight:bold;">' + r.contractor + '</td>'
      + '<td style="' + tdBase + '">' + r.date + '</td>'
      + '<td style="border:1px solid #ccc;padding:4px 5px;font-size:11px;white-space:normal;word-break:break-all;">' + r.company + ' ' + r.productName + '</td>'
      + '<td style="' + tdBase + 'text-align:right;">' + _fmtWon_(r.premium) + '원</td></tr>';
  }
  html += '</table><p style="color:#888;font-size:11px;margin-top:16px;">본 메일은 헤리티지 프로에서 자동 발송되었습니다.</p></div>';
  return html;
}


// 트리거 설치: 매일 오전 8시 (한 번만 실행)
function installAnniversaryTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendAnniversaryReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 매주 월요일 오전 8시 발송
  ScriptApp.newTrigger('sendAnniversaryReminders')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .inTimezone('Asia/Seoul')
    .create();
  return '매주 월요일 오전 8시 알림 트리거가 설치되었습니다.';
}

// 테스트용: 지금 즉시 실행 (메일 실제 발송됨)
function testAnniversaryNow() { return sendAnniversaryReminders(); }

// [권한 승인용] 이 함수를 한 번 실행하면 트리거·메일 권한 동의 창이 함께 뜬다.
// 실제로 아무것도 바꾸지 않고 권한만 요청한다. (승인 후 installAnniversaryTrigger 실행)
function authorizeAllPermissions() {
  var triggers = ScriptApp.getProjectTriggers();       // 트리거 권한(script.scriptapp) 유발
  var triggerCount = triggers.length;
  var remainingQuota = MailApp.getRemainingDailyQuota(); // 메일 권한(send_mail) 유발
  var ssName = '';
  try { ssName = SpreadsheetApp.openByUrl(SHEET_URL).getName(); } catch (e) { ssName = '(열기 실패)'; }
  return '권한 확인 완료 | 기존 트리거 ' + triggerCount + '개 | 오늘 남은 메일 ' + remainingQuota + '통 | 시트: ' + ssName;
}


// ============================================================
// 고객 관리 시트 내보내기
// 헤리티지 프로의 모든 청약을 FC 이름별 탭에 저장
// 대상 문서: 1VSAyjDxU8XughT4mo9V_B_JYd_9j_dcKlQZMywa5w4g
// 컬럼: 이름 연락처 주소 생년월일 가입상품 월보험료 상령일 상태 가족관계 최근접촉일 메모 위도 경도
// 방식: 계약마다 한 줄, 각 탭 기존 내용은 지우고 새로 씀(덮어쓰기)
// ============================================================
var CUSTOMER_EXPORT_SHEET_ID = '1VSAyjDxU8XughT4mo9V_B_JYd_9j_dcKlQZMywa5w4g';
var CUSTOMER_EXPORT_HEADERS = ['이름', '연락처', '주소', '생년월일', '가입상품', '월보험료', '상령일', '상태', '가족관계', '최근접촉일', '메모', '위도', '경도'];

// 본인계약 판별: 본인계약 체크 시 isBonusExcluded=true (실손도 같은 플래그라 isSilson은 제외)
function _isSelfContract_(it) {
  var a = (it && it.attributes) ? it.attributes : {};
  return !!a.isBonusExcluded && !a.isSilson;
}

function exportCustomersToSheet() {
  var ss = SpreadsheetApp.openById(CUSTOMER_EXPORT_SHEET_ID);
  // 문서의 기존 탭 이름 집합
  var existingSheets = {};
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    existingSheets[sheets[s].getName()] = sheets[s];
  }

  // 헤리티지 프로 전체 청약 로드 → FC별 그룹핑
  var rows = getAllData(); // [{name, month, json}]
  var byFc = {};
  for (var i = 0; i < rows.length; i++) {
    var fc = String(rows[i].name || '').trim();
    if (!fc) continue;
    var json = rows[i].json;
    if (!json) continue;
    var items;
    try { items = JSON.parse(json); } catch (e) { continue; }
    if (!items || !items.length) continue;
    for (var k = 0; k < items.length; k++) {
      var it = items[k];
      if (!it || it.type || !it.productKey || it.isWithdrawn || it.isSharedView) continue;
      if (_isSelfContract_(it)) continue; // 본인계약 제외
      if (!byFc[fc]) byFc[fc] = [];
      byFc[fc].push([
        it.contractor || '',                              // 이름(고객명)
        '',                                               // 연락처
        '',                                               // 주소
        '',                                               // 생년월일
        ((it.company || '') + ' ' + (it.productName || '')).trim(), // 가입상품
        Number(it.premium) || 0,                          // 월보험료
        '',                                               // 상령일
        '',                                               // 상태
        '',                                               // 가족관계
        '',                                               // 최근접촉일
        it.memo || '',                                    // 메모
        '',                                               // 위도
        ''                                                // 경도
      ]);
    }
  }

  var written = 0, tabsWritten = 0, skipped = [];
  for (var fcName in byFc) {
    var sheet = existingSheets[fcName];
    if (!sheet) { skipped.push(fcName); continue; } // 해당 이름 탭이 없으면 건너뜀
    var data = byFc[fcName];

    // 기존 내용 전체 삭제 후 헤더 + 데이터 새로 씀
    sheet.clearContents();
    sheet.getRange(1, 1, 1, CUSTOMER_EXPORT_HEADERS.length).setValues([CUSTOMER_EXPORT_HEADERS]);
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, CUSTOMER_EXPORT_HEADERS.length).setValues(data);
    }
    written += data.length;
    tabsWritten++;
  }

  var msg = '저장 완료: ' + tabsWritten + '개 탭, 총 ' + written + '건';
  if (skipped.length) msg += ' | 탭 없어 건너뜀: ' + skipped.join(', ');
  return msg;
}

// 특정 FC 한 명의 청약만 해당 탭에 갱신 (청약 저장 시 자동 호출용)
// createIfMissing=true 면 탭이 없을 때 새로 만든다.
function exportOneFcToSheet(fcName, createIfMissing) {
  fcName = String(fcName || '').trim();
  if (!fcName) return false;
  try {
    var ss = SpreadsheetApp.openById(CUSTOMER_EXPORT_SHEET_ID);
    var sheet = ss.getSheetByName(fcName);
    if (!sheet) {
      if (createIfMissing) {
        sheet = ss.insertSheet(fcName);
      } else {
        return false; // 탭 없으면 스킵
      }
    }

    // 이 FC의 모든 청약 수집 (전체 월)
    var rows = getAllData();
    var data = [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].name || '').trim() !== fcName) continue;
      if (!rows[i].json) continue;
      var items;
      try { items = JSON.parse(rows[i].json); } catch (e) { continue; }
      if (!items || !items.length) continue;
      for (var k = 0; k < items.length; k++) {
        var it = items[k];
        if (!it || it.type || !it.productKey || it.isWithdrawn || it.isSharedView) continue;
        if (_isSelfContract_(it)) continue; // 본인계약 제외
        data.push([
          it.contractor || '', '', '', '',
          ((it.company || '') + ' ' + (it.productName || '')).trim(),
          Number(it.premium) || 0, '', '', '', '',
          it.memo || '', '', ''
        ]);
      }
    }

    sheet.clearContents();
    sheet.getRange(1, 1, 1, CUSTOMER_EXPORT_HEADERS.length).setValues([CUSTOMER_EXPORT_HEADERS]);
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, CUSTOMER_EXPORT_HEADERS.length).setValues(data);
    }
    return true;
  } catch (e) {
    return false;
  }
}


// ====================================================================================
// [연체·실효 관리]
// 총무가 올린 보험사 취합 '환수예상 리스트'를 고객관리 시트(청약 내보내기와 같은 문서)에 저장하고,
// 관리자가 건별로 FC 면담 내용을 남긴다.
//
// 저장 탭 (고객관리 시트 문서 = CUSTOMER_EXPORT_SHEET_ID)
//   · 연체실효관리      : 건별 현황 (증권번호가 고유키)
//   · 연체실효_면담이력  : 면담 기록 (계속 누적, 수정·삭제 없음)
//
// 권한 — 서버에서 '사용자 정보'·'조직도' 시트로 직접 판정한다 (화면에서 넘긴 직책은 믿지 않음)
//   · 지사장 : 전체 조회 + 업로드 + 건 수정·삭제 + 면담 기록
//   · 총무   : 전체 조회 + 업로드 + 건 수정·삭제 (면담 기록은 하지 않음)
//   · 본부장 : 조직도 '세미나팀' 소속 FC + 본인 건 / 면담 기록
//   · 지점장 : 조직도에서 같은 하위소속(지점) FC + 본인 건 / 면담 기록
//   · 그 외(FC) : 본인 건만 조회 (관리자 면담 내용은 보이지 않음)
// ====================================================================================
var LAPSE_TAB = '연체실효관리';
var LAPSE_LOG_TAB = '연체실효_면담이력';
var LAPSE_HEADERS = ['관리ID', '최초등록월', '최근리스트월', '보험사', 'FC', '엑셀지점', '증권번호', '계약자', '상품명',
                     '월보험료', '계약일자', '최종납회', '유형', '예상환수액', '상태', '최근면담일', '다음확인일',
                     '최근면담자', '업로드일시', '업로드자', '수정일시', '수정자'];
var LAPSE_LOG_HEADERS = ['이력ID', '관리ID', '면담일', '작성자', '작성자직책', '고객사유', '면담내용', '조치계획',
                         '다음확인일', '결과상태', '작성일시'];
// 증권번호·날짜 등은 숫자/날짜로 자동 변환되지 않도록 '일반 텍스트' 서식을 고정한다.
var LAPSE_TEXT_COLS = ['관리ID', '최초등록월', '최근리스트월', '증권번호', '계약일자', '최근면담일', '다음확인일', '업로드일시', '수정일시'];
var LAPSE_LOG_TEXT_COLS = ['이력ID', '관리ID', '면담일', '고객사유', '면담내용', '조치계획', '다음확인일', '작성일시'];
var LAPSE_STATUSES = ['미확인', '추적중', '정상화', '해지확정'];
var LAPSE_MASK_NAMES = false;  // true 로 바꾸면 풀네임 계약자의 가운데 글자를 가려서 저장 (김길동 → 김*동). 현재는 엑셀에 온 그대로 저장.

function _lapseNow_(fmt) {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', fmt || 'yyyy-MM-dd HH:mm:ss');
}

function _lapseStr_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  return String(v).trim();
}

function _lapseNum_(v) {
  var n = Number(String(v === null || v === undefined ? '' : v).replace(/,/g, ''));
  return isFinite(n) ? n : 0;
}

// 수식으로 해석될 수 있는 맨 앞 글자(= + - @)는 공백을 앞에 붙여 무력화
function _lapseSafe_(s) {
  s = String(s === null || s === undefined ? '' : s).trim();
  return /^[=+\-@]/.test(s) ? ' ' + s : s;
}

function _lapseDate_(s) {
  s = _lapseStr_(s);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function _lapseMaskName_(nm) {
  nm = String(nm || '').trim();
  if (!LAPSE_MASK_NAMES || !nm || nm.indexOf('*') >= 0) return nm;
  if (!/^[가-힣]{2,5}$/.test(nm)) return nm;
  if (nm.length === 2) return nm.charAt(0) + '*';
  return nm.charAt(0) + new Array(nm.length - 1).join('*') + nm.charAt(nm.length - 1);
}

// 탭 서식: 일반 텍스트 열 고정
function _lapseFormat_(sh, hdr, textCols) {
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  for (var i = 0; i < textCols.length; i++) {
    var c = hdr.indexOf(textCols[i]);
    if (c >= 0) sh.getRange(2, c + 1, rows, 1).setNumberFormat('@');
  }
}

// 탭 확보: 없으면 생성, 헤더 보정(빠진 열은 오른쪽 끝에 추가), 서식 적용. 최종 헤더 배열 반환용으로 sh 와 hdr 를 함께 돌려준다.
function _lapseEnsureTab_(ss, name, headers, textCols) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  while (hdr.length && !hdr[hdr.length - 1]) hdr.pop();
  var fresh = (hdr.length === 0);
  headers.forEach(function (h) { if (hdr.indexOf(h) < 0) hdr.push(h); });
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  if (fresh) {
    sh.getRange(1, 1, 1, hdr.length).setFontWeight('bold').setBackground('#eef2ff');
    sh.setFrozenRows(1);
  }
  _lapseFormat_(sh, hdr, textCols);
  return { sh: sh, hdr: hdr };
}

// 탭 전체를 객체 배열로 (헤더 이름 기준)
function _lapseReadTab_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getDataRange().getValues();
  var hdr = vals[0].map(function (h) { return String(h).trim(); });
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    var o = { _row: i + 1 };
    var any = false;
    for (var c = 0; c < hdr.length; c++) {
      if (!hdr[c]) continue;
      o[hdr[c]] = vals[i][c];
      if (vals[i][c] !== '' && vals[i][c] !== null) any = true;
    }
    if (any) out.push(o);
  }
  return out;
}

// 호출자 권한·열람 범위 판정 (서버가 사용자 정보/조직도 시트로 직접 확인)
function _lapseCtx_(who) {
  who = String(who || '').trim();
  if (!who) return null;
  var users = getUserInfo(_lapseNow_('yyyy-MM')) || [];
  var me = null, userSet = {};
  for (var i = 0; i < users.length; i++) {
    userSet[users[i].name] = true;
    if (users[i].name === who) me = users[i];
  }
  if (!me) return null;

  var role = String(me.role || '');
  var isAll = role.indexOf('지사장') >= 0 || role.indexOf('총무') >= 0;
  var isHq = !isAll && role.indexOf('본부장') >= 0;
  var isBm = !isAll && !isHq && role.indexOf('지점장') >= 0;

  var org = [];
  try { org = getOrgChartData() || []; } catch (e) { org = []; }
  var norm = function (s) { return String(s || '').replace(/\s/g, ''); };

  var names = {};
  names[who] = true;                       // 본인 건은 항상 포함
  var label = '본인 건';
  if (isAll) {
    label = '전체';
  } else if (isHq) {
    org.forEach(function (m) { if (norm(m.subDept).indexOf('세미나') >= 0) names[m.name] = true; });
    label = '세미나팀';
  } else if (isBm) {
    var mySub = '';
    org.forEach(function (m) { if (m.name === who && norm(m.subDept)) mySub = norm(m.subDept); });
    if (mySub) {
      org.forEach(function (m) { if (norm(m.subDept) === mySub) names[m.name] = true; });
      label = mySub;
    }
  }
  return {
    name: who, role: role, all: isAll, names: names, userSet: userSet, scopeLabel: label,
    canUpload: isAll,
    canEdit: isAll,
    canWrite: (role.indexOf('지사장') >= 0) || isHq || isBm,
    canReadNotes: isAll || isHq || isBm
  };
}

function _lapseInScope_(ctx, fc) {
  return ctx.all || ctx.names[_lapseStr_(fc)] === true;
}

// ------------------------------------------------------------------
// 업로드 (지사장·총무): 환수예상 리스트 → '연체실효관리' 탭. 증권번호 기준으로 신규 추가/갱신.
//   rows: [{company, fc, branch, policyNo, contractor, product, premium, contractDate, lastPay, type, refund}]
// ------------------------------------------------------------------
function lapseUploadList(who, month, rows) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var ctx = _lapseCtx_(who);
    if (!ctx || !ctx.canUpload) return { success: false, error: '업로드 권한이 없습니다. (지사장·총무 전용)' };
    month = String(month || '').trim();
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) return { success: false, error: '리스트 기준월(YYYY-MM)을 확인해 주세요.' };
    if (!rows || !rows.length) return { success: false, error: '업로드할 데이터가 없습니다.' };
    if (rows.length > 3000) return { success: false, error: '한 번에 3,000건까지 업로드할 수 있습니다.' };

    lock.waitLock(20000);
    locked = true;

    var ss = SpreadsheetApp.openById(CUSTOMER_EXPORT_SHEET_ID);
    var t = _lapseEnsureTab_(ss, LAPSE_TAB, LAPSE_HEADERS, LAPSE_TEXT_COLS);
    var sh = t.sh, hdr = t.hdr;
    var C = {};
    hdr.forEach(function (h, i) { C[h] = i; });

    var data = [];
    if (sh.getLastRow() >= 2) data = sh.getRange(2, 1, sh.getLastRow() - 1, hdr.length).getValues();
    var byId = {};
    data.forEach(function (r, i) { var k = _lapseStr_(r[C['관리ID']]); if (k) byId[k] = i; });

    var now = _lapseNow_();
    var added = 0, updated = 0, unchanged = 0, reopened = 0, skipped = 0;
    var unmatched = {};

    for (var n = 0; n < rows.length; n++) {
      var r = rows[n] || {};
      var id = String(r.policyNo === null || r.policyNo === undefined ? '' : r.policyNo).replace(/\s/g, '');
      var fc = String(r.fc || '').trim();
      if (!id || !fc) { skipped++; continue; }
      if (!ctx.userSet[fc]) unmatched[fc] = true;

      var rec = {
        '보험사': _lapseSafe_(r.company),
        'FC': fc,
        '엑셀지점': _lapseSafe_(r.branch),
        '증권번호': id,
        '계약자': _lapseSafe_(_lapseMaskName_(r.contractor)),
        '상품명': _lapseSafe_(r.product),
        '월보험료': _lapseNum_(r.premium),
        '계약일자': _lapseDate_(r.contractDate),
        '최종납회': _lapseNum_(r.lastPay),
        '유형': _lapseSafe_(r.type) || '미분류',
        '예상환수액': _lapseNum_(r.refund)
      };

      if (byId.hasOwnProperty(id)) {
        var row = data[byId[id]];
        var changed = (_lapseStr_(row[C['유형']]) !== rec['유형']) ||
                      (_lapseNum_(row[C['예상환수액']]) !== rec['예상환수액']) ||
                      (_lapseStr_(row[C['FC']]) !== rec['FC']) ||
                      (_lapseStr_(row[C['계약자']]) !== rec['계약자']);
        Object.keys(rec).forEach(function (k) { row[C[k]] = rec[k]; });
        var prevMonth = _lapseStr_(row[C['최근리스트월']]);
        if (!prevMonth || month > prevMonth) {
          row[C['최근리스트월']] = month;
          // 정상화됐던 건이 더 최신 리스트에 다시 올라오면 다시 확인 대상으로 되돌린다
          if (_lapseStr_(row[C['상태']]) === '정상화') { row[C['상태']] = '미확인'; reopened++; changed = true; }
        }
        row[C['업로드일시']] = now;
        row[C['업로드자']] = ctx.name;
        if (changed) updated++; else unchanged++;
      } else {
        var nr = new Array(hdr.length);
        for (var z = 0; z < nr.length; z++) nr[z] = '';
        Object.keys(rec).forEach(function (k) { nr[C[k]] = rec[k]; });
        nr[C['관리ID']] = id;
        nr[C['최초등록월']] = month;
        nr[C['최근리스트월']] = month;
        nr[C['상태']] = '미확인';
        nr[C['업로드일시']] = now;
        nr[C['업로드자']] = ctx.name;
        data.push(nr);
        byId[id] = data.length - 1;
        added++;
      }
    }

    // 행이 모자라면 늘리고 서식 다시 적용
    var need = data.length + 1;
    if (sh.getMaxRows() < need) {
      sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
      _lapseFormat_(sh, hdr, LAPSE_TEXT_COLS);
    }
    if (data.length) sh.getRange(2, 1, data.length, hdr.length).setValues(data);

    return {
      success: true, total: rows.length, added: added, updated: updated, unchanged: unchanged,
      reopened: reopened, skipped: skipped, unmatched: Object.keys(unmatched)
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e2) {} }
  }
}

// ------------------------------------------------------------------
// 목록 조회: 호출자의 열람 범위에 해당하는 건만 반환
// ------------------------------------------------------------------
function lapseGetList(who) {
  try {
    var ctx = _lapseCtx_(who);
    if (!ctx) return { success: false, error: '사용자 정보를 확인할 수 없습니다.' };
    var ss = SpreadsheetApp.openById(CUSTOMER_EXPORT_SHEET_ID);
    var main = _lapseReadTab_(ss.getSheetByName(LAPSE_TAB));

    var cnt = {};
    if (ctx.canReadNotes) {
      _lapseReadTab_(ss.getSheetByName(LAPSE_LOG_TAB)).forEach(function (l) {
        var k = _lapseStr_(l['관리ID']);
        if (k) cnt[k] = (cnt[k] || 0) + 1;
      });
    }

    var out = [];
    main.forEach(function (r) {
      if (!_lapseInScope_(ctx, r['FC'])) return;
      var id = _lapseStr_(r['관리ID']);
      if (!id) return;
      var item = {
        id: id,
        month: _lapseStr_(r['최근리스트월']),
        firstMonth: _lapseStr_(r['최초등록월']),
        company: _lapseStr_(r['보험사']),
        fc: _lapseStr_(r['FC']),
        branch: _lapseStr_(r['엑셀지점']),
        policyNo: _lapseStr_(r['증권번호']),
        contractor: _lapseStr_(r['계약자']),
        product: _lapseStr_(r['상품명']),
        premium: _lapseNum_(r['월보험료']),
        contractDate: _lapseStr_(r['계약일자']),
        lastPay: _lapseNum_(r['최종납회']),
        type: _lapseStr_(r['유형']),
        refund: _lapseNum_(r['예상환수액']),
        status: _lapseStr_(r['상태']) || '미확인',
        lastTalk: _lapseStr_(r['최근면담일']),
        nextCheck: _lapseStr_(r['다음확인일']),
        talkCount: cnt[id] || 0
      };
      if (ctx.canReadNotes) item.lastTalker = _lapseStr_(r['최근면담자']);
      out.push(item);
    });

    return {
      success: true,
      me: { name: ctx.name, role: ctx.role, scopeLabel: ctx.scopeLabel, canUpload: ctx.canUpload, canEdit: ctx.canEdit,
            canWrite: ctx.canWrite, canReadNotes: ctx.canReadNotes },
      rows: out
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ------------------------------------------------------------------
// 면담 이력 조회 (지사장·총무·본부장·지점장 — 열람 범위 안의 건만)
// ------------------------------------------------------------------
function lapseGetHistory(who, id) {
  try {
    var ctx = _lapseCtx_(who);
    if (!ctx || !ctx.canReadNotes) return { success: false, error: '면담 내용을 볼 수 있는 권한이 없습니다.' };
    id = String(id || '').trim();
    var ss = SpreadsheetApp.openById(CUSTOMER_EXPORT_SHEET_ID);
    var found = null;
    _lapseReadTab_(ss.getSheetByName(LAPSE_TAB)).forEach(function (r) {
      if (_lapseStr_(r['관리ID']) === id) found = r;
    });
    if (!found || !_lapseInScope_(ctx, found['FC'])) return { success: false, error: '해당 건을 찾을 수 없습니다.' };

    var list = [];
    _lapseReadTab_(ss.getSheetByName(LAPSE_LOG_TAB)).forEach(function (l) {
      if (_lapseStr_(l['관리ID']) !== id) return;
      list.push({
        talkDate: _lapseStr_(l['면담일']), writer: _lapseStr_(l['작성자']), writerRole: _lapseStr_(l['작성자직책']),
        reason: _lapseStr_(l['고객사유']), content: _lapseStr_(l['면담내용']), plan: _lapseStr_(l['조치계획']),
        nextCheck: _lapseStr_(l['다음확인일']), result: _lapseStr_(l['결과상태']), createdAt: _lapseStr_(l['작성일시'])
      });
    });
    list.sort(function (a, b) {
      if (a.talkDate !== b.talkDate) return a.talkDate < b.talkDate ? 1 : -1;
      return a.createdAt < b.createdAt ? 1 : -1;
    });
    return { success: true, history: list };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ------------------------------------------------------------------
// 면담 기록 저장 (지사장·본부장·지점장 — 열람 범위 안의 건만)
//   p: {id, talkDate, reason, content, plan, result, nextCheck}
// ------------------------------------------------------------------
function lapseSaveInterview(who, p) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var ctx = _lapseCtx_(who);
    if (!ctx || !ctx.canWrite) return { success: false, error: '면담 기록 권한이 없습니다.' };
    p = p || {};
    var id = String(p.id || '').trim();
    var talkDate = _lapseDate_(p.talkDate);
    var result = String(p.result || '').trim();
    var content = String(p.content || '').trim();
    var nextCheck = _lapseDate_(p.nextCheck);
    if (!id) return { success: false, error: '대상 건이 지정되지 않았습니다.' };
    if (!talkDate) return { success: false, error: '면담일을 입력해 주세요.' };
    if (LAPSE_STATUSES.indexOf(result) < 1) return { success: false, error: '결과 상태를 선택해 주세요.' };
    if (content.length < 5) return { success: false, error: '면담 내용을 5자 이상 입력해 주세요.' };
    if (content.length > 3000) return { success: false, error: '면담 내용은 3,000자 이내로 입력해 주세요.' };
    if (result === '추적중' && !nextCheck) return { success: false, error: '추적중으로 두려면 다음 확인일을 입력해 주세요.' };

    lock.waitLock(20000);
    locked = true;

    var ss = SpreadsheetApp.openById(CUSTOMER_EXPORT_SHEET_ID);
    var mainTab = _lapseEnsureTab_(ss, LAPSE_TAB, LAPSE_HEADERS, LAPSE_TEXT_COLS);
    var sh = mainTab.sh, hdr = mainTab.hdr;
    var C = {};
    hdr.forEach(function (h, i) { C[h] = i; });

    var rowNo = -1;
    if (sh.getLastRow() >= 2) {
      var ids = sh.getRange(2, C['관리ID'] + 1, sh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (_lapseStr_(ids[i][0]) === id) { rowNo = i + 2; break; }
      }
    }
    if (rowNo < 0) return { success: false, error: '해당 건을 찾을 수 없습니다.' };
    var cur = sh.getRange(rowNo, 1, 1, hdr.length).getValues()[0];
    if (!_lapseInScope_(ctx, cur[C['FC']])) return { success: false, error: '열람 범위 밖의 건입니다.' };

    // 이력 한 줄 추가 (누적 기록)
    var lt = _lapseEnsureTab_(ss, LAPSE_LOG_TAB, LAPSE_LOG_HEADERS, LAPSE_LOG_TEXT_COLS);
    var L = {};
    lt.hdr.forEach(function (h, i) { L[h] = i; });
    var logRow = new Array(lt.hdr.length);
    for (var z = 0; z < logRow.length; z++) logRow[z] = '';
    logRow[L['이력ID']] = Utilities.getUuid().substring(0, 8);
    logRow[L['관리ID']] = id;
    logRow[L['면담일']] = talkDate;
    logRow[L['작성자']] = ctx.name;
    logRow[L['작성자직책']] = ctx.role;
    logRow[L['고객사유']] = _lapseSafe_(p.reason).substring(0, 100);
    logRow[L['면담내용']] = _lapseSafe_(content);
    logRow[L['조치계획']] = _lapseSafe_(p.plan).substring(0, 300);
    logRow[L['다음확인일']] = nextCheck;
    logRow[L['결과상태']] = result;
    logRow[L['작성일시']] = _lapseNow_();
    var logNo = lt.sh.getLastRow() + 1;
    if (lt.sh.getMaxRows() < logNo) {
      lt.sh.insertRowsAfter(lt.sh.getMaxRows(), 100);
      _lapseFormat_(lt.sh, lt.hdr, LAPSE_LOG_TEXT_COLS);
    }
    lt.sh.getRange(logNo, 1, 1, lt.hdr.length).setValues([logRow]);

    // 건별 현황 갱신
    cur[C['상태']] = result;
    cur[C['최근면담일']] = talkDate;
    cur[C['다음확인일']] = nextCheck;
    cur[C['최근면담자']] = ctx.name;
    sh.getRange(rowNo, 1, 1, hdr.length).setValues([cur]);

    return { success: true, status: result, lastTalk: talkDate, nextCheck: nextCheck };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e2) {} }
  }
}
// ------------------------------------------------------------------
// 건 수정 (지사장·총무)
//   증권번호(고유키)·상태·면담 기록은 그대로 두고 나머지 항목을 고친다. 수정일시·수정자를 남긴다.
//   p: {company, fc, branch, contractor, product, premium, contractDate, lastPay, type, refund}
// ------------------------------------------------------------------
function lapseUpdateCase(who, id, p) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var ctx = _lapseCtx_(who);
    if (!ctx || !ctx.canEdit) return { success: false, error: '수정 권한이 없습니다. (지사장·총무 전용)' };
    id = String(id || '').trim();
    p = p || {};
    var fc = String(p.fc || '').trim();
    var type = _lapseSafe_(p.type);
    var cd = String(p.contractDate || '').trim();
    if (!id) return { success: false, error: '대상 건이 지정되지 않았습니다.' };
    if (!fc) return { success: false, error: 'FC 이름을 입력해 주세요.' };
    if (!type) return { success: false, error: '유형(실효·해약·연체 등)을 입력해 주세요.' };
    if (cd && !_lapseDate_(cd)) return { success: false, error: '계약일자 형식이 올바르지 않습니다.' };

    lock.waitLock(20000);
    locked = true;

    var ss = SpreadsheetApp.openById(CUSTOMER_EXPORT_SHEET_ID);
    var t = _lapseEnsureTab_(ss, LAPSE_TAB, LAPSE_HEADERS, LAPSE_TEXT_COLS);
    var sh = t.sh, hdr = t.hdr;
    var C = {};
    hdr.forEach(function (h, i) { C[h] = i; });

    var rowNo = -1;
    if (sh.getLastRow() >= 2) {
      var ids = sh.getRange(2, C['관리ID'] + 1, sh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (_lapseStr_(ids[i][0]) === id) { rowNo = i + 2; break; }
      }
    }
    if (rowNo < 0) return { success: false, error: '해당 건을 찾을 수 없습니다. (이미 삭제되었을 수 있습니다)' };

    var cur = sh.getRange(rowNo, 1, 1, hdr.length).getValues()[0];
    cur[C['보험사']] = _lapseSafe_(p.company);
    cur[C['FC']] = fc;
    cur[C['엑셀지점']] = _lapseSafe_(p.branch);
    cur[C['계약자']] = _lapseSafe_(p.contractor);
    cur[C['상품명']] = _lapseSafe_(p.product);
    cur[C['월보험료']] = _lapseNum_(p.premium);
    cur[C['계약일자']] = cd;
    cur[C['최종납회']] = _lapseNum_(p.lastPay);
    cur[C['유형']] = type;
    cur[C['예상환수액']] = _lapseNum_(p.refund);
    cur[C['수정일시']] = _lapseNow_();
    cur[C['수정자']] = ctx.name;
    sh.getRange(rowNo, 1, 1, hdr.length).setValues([cur]);

    return { success: true, unmatchedFc: !ctx.userSet[fc] };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e2) {} }
  }
}

// ------------------------------------------------------------------
// 건 삭제 (지사장·총무): 현황 행과 그 건의 면담 이력을 함께 지운다. 되돌릴 수 없다.
// ------------------------------------------------------------------
function lapseDeleteCase(who, id) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var ctx = _lapseCtx_(who);
    if (!ctx || !ctx.canEdit) return { success: false, error: '삭제 권한이 없습니다. (지사장·총무 전용)' };
    id = String(id || '').trim();
    if (!id) return { success: false, error: '대상 건이 지정되지 않았습니다.' };

    lock.waitLock(20000);
    locked = true;

    var ss = SpreadsheetApp.openById(CUSTOMER_EXPORT_SHEET_ID);
    var sh = ss.getSheetByName(LAPSE_TAB);
    if (!sh || sh.getLastRow() < 2) return { success: false, error: '해당 건을 찾을 수 없습니다.' };
    var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    var idCol = hdr.indexOf('관리ID');
    if (idCol < 0) return { success: false, error: '시트 형식을 확인할 수 없습니다.' };
    var ids = sh.getRange(2, idCol + 1, sh.getLastRow() - 1, 1).getValues();
    var rowNo = -1;
    for (var i = 0; i < ids.length; i++) {
      if (_lapseStr_(ids[i][0]) === id) { rowNo = i + 2; break; }
    }
    if (rowNo < 0) return { success: false, error: '해당 건을 찾을 수 없습니다. (이미 삭제되었을 수 있습니다)' };
    sh.deleteRow(rowNo);

    // 면담 이력도 함께 삭제 (아래에서 위로 지워야 행 번호가 밀리지 않는다)
    var removed = 0;
    var logSh = ss.getSheetByName(LAPSE_LOG_TAB);
    if (logSh && logSh.getLastRow() >= 2) {
      var lh = logSh.getRange(1, 1, 1, logSh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
      var lc = lh.indexOf('관리ID');
      if (lc >= 0) {
        var lv = logSh.getRange(2, lc + 1, logSh.getLastRow() - 1, 1).getValues();
        for (var j = lv.length - 1; j >= 0; j--) {
          if (_lapseStr_(lv[j][0]) === id) { logSh.deleteRow(j + 2); removed++; }
        }
      }
    }
    return { success: true, removedLogs: removed };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e2) {} }
  }
}
// ===== [연체·실효 관리] 끝 =====


// ===================== [브리핑 계약] 마스터본부 탭 =====================
// 고객관리 시트 문서(CUSTOMER_EXPORT_SHEET_ID)의 '마스터본부' 탭: A=섭외자 이름, B=쉐어율(20% 또는 0.2 또는 20)
var MASTERHQ_TAB = '마스터본부';
function _mhqRate_(v) {
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).trim();
  var isPct = s.indexOf('%') >= 0;
  var n = parseFloat(s.replace(/[%,\s]/g, ''));
  if (isNaN(n)) return 0;
  if (isPct) return n / 100;
  return n > 1 ? n / 100 : n;   // 0.2 → 20%, 20 → 20%
}
function getMasterHqBrokers() {
  try {
    var ss = SpreadsheetApp.openById(CUSTOMER_EXPORT_SHEET_ID);
    var sh = ss.getSheetByName(MASTERHQ_TAB);
    if (!sh) {
      sh = ss.insertSheet(MASTERHQ_TAB);
      sh.getRange(1, 1, 1, 2).setValues([['섭외자 이름', '쉐어율']]);
      sh.getRange(1, 1, 1, 2).setFontWeight('bold');
      sh.getRange(2, 2, 500, 1).setNumberFormat('0%');
      return [];
    }
    var last = sh.getLastRow();
    if (last < 2) return [];
    var vals = sh.getRange(2, 1, last - 1, 2).getValues();
    var out = [], seen = {};
    vals.forEach(function (r) {
      var name = String(r[0] || '').trim();
      if (!name || seen[name]) return;
      seen[name] = 1;
      out.push({ name: name, rate: _mhqRate_(r[1]) });
    });
    return out;
  } catch (e) {
    return { error: String(e && e.message || e) };
  }
}
