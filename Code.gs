// =================================================================
// 2기계1반 학급시스템 - 초고속 캐시 최적화 백엔드 (Code.gs)
// =================================================================

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;
  const todayStrParam = e.parameter.todayStr;
  let result = { success: false, message: 'Invalid Action' };

  try {
    if (action === 'getInitialData') {
      result = getInitialDataFast(todayStrParam);
    } else if (action === 'getTodayAttendanceStatus') {
      result = getTodayAttendanceStatus(todayStrParam);
    } else if (action === 'getStudentAttendanceHistory') {
      result = getStudentAttendanceHistory(e.parameter.studentId);
    } else if (action === 'getMonthAttendanceMatrix') {
      result = getMonthAttendanceMatrix(e.parameter.yearMonth);
    } else {
      result = { success: true, status: 'Attendance Fast API Server is Running' };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  if (callback) {
    const jsonpOutput = callback + '(' + JSON.stringify(result) + ');';
    return ContentService.createTextOutput(jsonpOutput)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result = { success: false, message: 'Invalid Action' };
  try {
    let postData = {};
    if (e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
      } catch (ex) {
        postData = e.parameter || {};
      }
    } else {
      postData = e.parameter || {};
    }

    const action = postData.action;

    if (action === 'verifyTeacherPassword') {
      result = verifyTeacherPassword(postData.password);
    } else if (action === 'generateDailyPin') {
      result = generateDailyPin();
      invalidateFastCache();
    } else if (action === 'submitSelfAttendance') {
      result = submitSelfAttendance(postData.studentId, postData.pin, postData.lat, postData.lng);
      invalidateFastCache();
    } else if (action === 'updateManualAttendance') {
      result = updateManualAttendance(postData.studentId, postData.status, postData.period, postData.reason, postData.targetDate);
      invalidateFastCache();
    } else if (action === 'saveNoticeAndDDayData') {
      result = saveNoticeAndDDayData(postData.payload);
      invalidateFastCache();
    } else if (action === 'saveTimeLockSettings') {
      result = saveTimeLockSettings(postData.startTime, postData.endTime, postData.radius);
      invalidateFastCache();
    } else if (action === 'saveMealData') {
      result = saveMealData(postData.meals);
      invalidateFastCache();
    } else if (action === 'saveNeisData') {
      result = saveNeisData(postData.period, postData.neisData);
      invalidateFastCache();
    } else if (action === 'saveOnlineReportData') {
      result = saveOnlineReportData(postData.reportList);
      invalidateFastCache();
    } else if (action === 'deleteOnlineReportRecord') {
      result = deleteOnlineReportRecord(postData.studentId, postData.dateStr);
      invalidateFastCache();
    } else if (action === 'clearAllOnlineReports') {
      result = clearAllOnlineReports();
      invalidateFastCache();
    } else if (action === 'updateStudentRoster') {
      result = updateStudentRoster(postData.rosterList);
      invalidateFastCache();
    } else if (action === 'updateTeacherPassword') {
      result = updateTeacherPassword(postData.currentPassword, postData.newPassword);
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------
// ⚡ 초고속 캐시 엔진 (Fast Cache Layer)
// -------------------------------------------------------------

function normalizePin(pin) {
  if (pin === null || pin === undefined) return '';
  let str = String(pin).replace(/[\s\u00A0\u200B\uFEFF\r\n\t'"`,]+/g, '').trim();
  // 전각 숫자를 반각 숫자로 변환 (０-９ -> 0-9)
  str = str.replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  // 소수점 및 소수점 이하 제거 (예: "1234.0" -> "1234")
  str = str.replace(/\.0+$/, '');
  // 숫자 외 공백 및 특수문자 제거
  return str.replace(/[^0-9]/g, '');
}

function invalidateFastCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('APP_INITIAL_DATA');
  cache.remove('BOARD_DATA');
  cache.remove('MEAL_DATA');
  try {
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    cache.remove('APP_INITIAL_DATA_' + todayStr);
  } catch (e) {}
}

function getInitialDataFast(targetTodayStr) {
  const kstToday = targetTodayStr || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const cacheKey = 'APP_INITIAL_DATA_' + kstToday;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  const rawData = getInitialData(kstToday);
  try {
    cache.put(cacheKey, JSON.stringify(rawData), 600); // 10분 캐싱
  } catch (e) {}
  return rawData;
}

// -------------------------------------------------------------
// 스프레드시트 관리 및 비즈니스 로직
// -------------------------------------------------------------

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let studentSheet = ss.getSheetByName('학생명단');
  if (!studentSheet) studentSheet = ss.insertSheet('학생명단');
  if (studentSheet.getLastRow() === 0) {
    const initialStudentData = [['번호', '이름', '개인식별번호']];
    for (let i = 1; i <= 20; i++) initialStudentData.push([i, `학생${i}`, '0000']);
    studentSheet.getRange(1, 1, initialStudentData.length, 3).setValues(initialStudentData);
  }

  let recordSheet = ss.getSheetByName('출결기록');
  if (!recordSheet) recordSheet = ss.insertSheet('출결기록');
  if (recordSheet.getLastRow() === 0) {
    recordSheet.appendRow(['기록일시', '일자', '번호', '이름', '상태', '교시', '사유', '위치인증여부']);
  } else {
    ensureRecordSchema(recordSheet);
  }

  let neisSheet = ss.getSheetByName('나이스누적');
  if (!neisSheet) neisSheet = ss.insertSheet('나이스누적');
  if (neisSheet.getLastRow() === 0) {
    neisSheet.appendRow(['번호', '이름', '결석_질병', '결석_미인정', '결석_기타', '결석_인정', '지각_질병', '지각_미인정', '지각_기타', '지각_인정', '조퇴_질병', '조퇴_미인정', '조퇴_기타', '조퇴_인정', '결과_질병', '결과_미인정', '결과_기타', '결과_인정', '업데이트일시', '집계기간']);
  }

  let mealSheet = ss.getSheetByName('급식식단');
  if (!mealSheet) mealSheet = ss.insertSheet('급식식단');
  if (mealSheet.getLastRow() === 0) mealSheet.appendRow(['일자', '식단메뉴', '영양칼로리', '업데이트일시']);

  let noticeSheet = ss.getSheetByName('알림장');
  if (!noticeSheet) noticeSheet = ss.insertSheet('알림장');
  if (noticeSheet.getLastRow() === 0) {
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    noticeSheet.appendRow(['구분', '항목/제목', '날짜/내용', '추가정보', '업데이트일시']);
    noticeSheet.appendRow(['TITLE', '메인제목', '오늘의 학급 알림장', '', nowStr]);
    noticeSheet.appendRow(['D_DAY', '2학기 1차 지필평가', '2026-10-12', '중간고사', nowStr]);
    noticeSheet.appendRow(['NOTICE', '오늘의 학급 안내사항', '1. 내일 1교시 실습실 이동\n2. 과제 제출 마감 확인', '일반', nowStr]);
  }

  let configSheet = ss.getSheetByName('설정');
  if (!configSheet) {
    configSheet = ss.insertSheet('설정');
    const targetConfig = [
      ['항목', '값'],
      ['나이스집계기간', '2026.03.02. - 2026.07.31.'],
      ['학교위도', '37.278698087555604'],
      ['학교경도', '127.45853065238362'],
      ['허용반경m', '250'],
      ['오늘출석핀', ''],
      ['출석시작시간', '08:00'],
      ['출석마감시간', '08:50'],
      ['교사비밀번호', '19650018']
    ];
    configSheet.getRange(1, 1, targetConfig.length, 2).setValues(targetConfig);
  } else {
    ensureConfigSchema(configSheet);
  }

  let onlineReportSheet = ss.getSheetByName('온라인신고서');
  if (!onlineReportSheet) {
    onlineReportSheet = ss.insertSheet('온라인신고서');
    onlineReportSheet.appendRow(['번호', '이름', '일자', '구분', '진행상태', '결재상태', '교시기간', '신청일자', '업데이트일시']);
  }

  return ss;
}

function ensureConfigSchema(configSheet) {
  const rows = configSheet.getDataRange().getValues();
  const keys = rows.map(r => String(r[0]).trim());
  if (!keys.includes('출석시작시간')) configSheet.appendRow(['출석시작시간', '08:00']);
  if (!keys.includes('출석마감시간')) configSheet.appendRow(['출석마감시간', '08:50']);
  if (!keys.includes('교사비밀번호')) configSheet.appendRow(['교사비밀번호', '19650018']);
}

function ensureRecordSchema(recordSheet) {
  const lastCol = recordSheet.getLastColumn();
  const lastRow = recordSheet.getLastRow();
  if (lastCol === 0 || lastRow === 0) return;
  const headers = recordSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const periodIndex = headers.indexOf('교시');
  if (periodIndex === -1) {
    recordSheet.insertColumnAfter(5);
    recordSheet.getRange(1, 6).setValue('교시');
    if (lastRow > 1) {
      const defaultPeriods = Array.from({ length: lastRow - 1 }, () => ['종일']);
      recordSheet.getRange(2, 6, lastRow - 1, 1).setValues(defaultPeriods);
    }
  }
}

function getStudentMasterData(ss) {
  const studentSheet = ss.getSheetByName('학생명단');
  if (!studentSheet || studentSheet.getLastRow() <= 1) return [];
  const data = studentSheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const id = parseInt(data[i][0], 10);
    if (!isNaN(id)) {
      const rawPin = data[i][2];
      const pinStr = (rawPin !== null && rawPin !== undefined) ? normalizePin(rawPin) : '0000';
      students.push({ 
        id: id, 
        name: String(data[i][1] || '').trim(), 
        pin: pinStr || '0000' 
      });
    }
  }
  return students;
}

function formatDateToCustomString(dateObj) {
  if (!dateObj) return '';
  const d = (dateObj instanceof Date) ? dateObj : new Date(dateObj);
  if (isNaN(d.getTime())) return String(dateObj);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  let hours = d.getHours();
  const minutes = String(d.getMinutes()).padStart(2, '0');
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12;
  return `${year}.${month}.${day} ${ampm} ${hours}:${minutes}`;
}

function getColumnIndexMap(headerRow) {
  const map = { time: 0, date: 1, id: 2, name: 3, status: 4, period: -1, reason: -1, auth: -1 };
  headerRow.forEach((h, i) => {
    const col = String(h).trim();
    if (col === '기록일시') map.time = i;
    else if (col === '일자') map.date = i;
    else if (col === '번호') map.id = i;
    else if (col === '이름') map.name = i;
    else if (col === '상태') map.status = i;
    else if (col === '교시') map.period = i;
    else if (col === '사유') map.reason = i;
    else if (col === '위치인증여부') map.auth = i;
  });
  return map;
}

function fetchOnlineReportsMap(ssInstance) {
  const ss = ssInstance || setupSpreadsheet();
  const sheet = ss.getSheetByName('온라인신고서');
  const reports = {};
  if (!sheet || sheet.getLastRow() <= 1) return reports;

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const sId = parseInt(rows[i][0], 10);
    const rawDate = rows[i][2];
    const dStr = (rawDate instanceof Date) 
      ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') 
      : String(rawDate).trim();

    if (!isNaN(sId) && dStr) {
      if (!reports[sId]) reports[sId] = {};
      reports[sId][dStr] = {
        name: String(rows[i][1] || '').trim(),
        category: String(rows[i][3] || '').trim(),
        processStatus: String(rows[i][4] || '').trim(),
        approvalStatus: String(rows[i][5] || '').trim(),
        period: String(rows[i][6] || '').trim(),
        applyDate: String(rows[i][7] || '').trim()
      };
    }
  }
  return reports;
}

function saveOnlineReportData(reportList) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '다른 작업 처리 중입니다.' }; }
  try {
    const ss = setupSpreadsheet();
    let sheet = ss.getSheetByName('온라인신고서');
    if (!sheet) sheet = ss.insertSheet('온라인신고서');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const existingReports = fetchOnlineReportsMap(ss);

    reportList.forEach(item => {
      const sId = parseInt(item.id, 10);
      const dStr = String(item.date).trim();
      if (!isNaN(sId) && dStr) {
        if (!existingReports[sId]) existingReports[sId] = {};
        existingReports[sId][dStr] = {
          name: item.name || '',
          category: item.category || '결석',
          processStatus: item.processStatus || '접수',
          approvalStatus: item.approvalStatus || '완결',
          period: item.period || '종일',
          applyDate: item.applyDate || ''
        };
      }
    });

    const students = getStudentMasterData(ss);
    const nameMap = {};
    students.forEach(st => { nameMap[st.id] = st.name; });

    const rowsToSave = [['번호', '이름', '일자', '구분', '진행상태', '결재상태', '교시기간', '신청일자', '업데이트일시']];
    const allIds = Object.keys(existingReports).sort((a, b) => Number(a) - Number(b));

    allIds.forEach(id => {
      const sId = Number(id);
      const stName = nameMap[sId] || `학생${sId}`;
      const dateKeys = Object.keys(existingReports[sId]).sort();
      dateKeys.forEach(dKey => {
        const r = existingReports[sId][dKey];
        rowsToSave.push([
          sId, r.name || stName, dKey, r.category || '결석',
          r.processStatus || '접수', r.approvalStatus || '완결',
          r.period || '종일', r.applyDate || '', nowStr
        ]);
      });
    });

    sheet.clearContents();
    sheet.getRange(1, 1, rowsToSave.length, 9).setValues(rowsToSave);
    SpreadsheetApp.flush();

    return { 
      success: true, 
      count: reportList.length, 
      reportsMap: existingReports, 
      message: `온라인 출결 신고서 ${reportList.length}건이 성공적으로 동기화되었습니다.` 
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteOnlineReportRecord(studentId, dateStr) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '처리 중 오류가 발생했습니다.' }; }
  try {
    const ss = setupSpreadsheet();
    const sheet = ss.getSheetByName('온라인신고서');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, message: '삭제할 내역이 없습니다.', reportsMap: {} };

    const sId = parseInt(studentId, 10);
    const targetDate = String(dateStr).trim();
    const rows = sheet.getDataRange().getValues();
    let foundRow = -1;

    for (let i = 1; i < rows.length; i++) {
      const rowId = parseInt(rows[i][0], 10);
      const rawDate = rows[i][2];
      const rDate = (rawDate instanceof Date) 
        ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') 
        : String(rawDate).trim();

      if (rowId === sId && rDate === targetDate) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow !== -1) {
      sheet.deleteRow(foundRow);
      SpreadsheetApp.flush();
    }

    const updatedReports = fetchOnlineReportsMap(ss);
    return { success: true, message: `${sId}번 학생의 [${targetDate}] 온라인 신고서 내역이 삭제되었습니다.`, reportsMap: updatedReports };
  } finally {
    lock.releaseLock();
  }
}

function clearAllOnlineReports() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '처리 중 오류가 발생했습니다.' }; }
  try {
    const ss = setupSpreadsheet();
    let sheet = ss.getSheetByName('온라인신고서');
    if (sheet) {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, 9).setValues([['번호', '이름', '일자', '구분', '진행상태', '결재상태', '교시기간', '신청일자', '업데이트일시']]);
      SpreadsheetApp.flush();
    }
    return { success: true, message: '모든 온라인 신고서 동기화 내역이 초기화되었습니다.', reportsMap: {} };
  } finally {
    lock.releaseLock();
  }
}

function fetchTodayAttendanceMap(todayStr, ssInstance) {
  const ss = ssInstance || setupSpreadsheet();
  const recordSheet = ss.getSheetByName('출결기록');
  if (!recordSheet || recordSheet.getLastRow() <= 1) return {};

  const recordRows = recordSheet.getDataRange().getValues();
  const colMap = getColumnIndexMap(recordRows[0]);
  const todayRecords = {};

  for (let i = recordRows.length - 1; i >= 1; i--) {
    const rawDate = recordRows[i][colMap.date];
    const rDate = (rawDate instanceof Date) 
      ? Utilities.formatDate(rawDate, 'Asia/Seoul', 'yyyy-MM-dd') 
      : String(rawDate).trim();

    if (rDate === todayStr) {
      const num = parseInt(recordRows[i][colMap.id], 10);
      if (!isNaN(num) && !todayRecords[num]) {
        todayRecords[num] = {
          time: formatDateToCustomString(recordRows[i][colMap.time]),
          status: String(recordRows[i][colMap.status] || '출석'),
          period: (colMap.period !== -1 && recordRows[i][colMap.period]) ? String(recordRows[i][colMap.period]) : '종일',
          reason: (colMap.reason !== -1 && recordRows[i][colMap.reason]) ? String(recordRows[i][colMap.reason]) : '',
          gpsAuth: (colMap.auth !== -1 && recordRows[i][colMap.auth]) ? String(recordRows[i][colMap.auth]) : '인증완료'
        };
      }
    }
  }
  return todayRecords;
}

function fetchMealData(ssInstance) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('MEAL_DATA');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const ss = ssInstance || setupSpreadsheet();
  const mealSheet = ss.getSheetByName('급식식단');
  const meals = {};
  if (!mealSheet || mealSheet.getLastRow() <= 1) return meals;

  const rows = mealSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rawDate = rows[i][0];
    const dStr = (rawDate instanceof Date) 
      ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') 
      : String(rawDate).trim();
    if (dStr) {
      meals[dStr] = { menu: String(rows[i][1] || '').trim(), nutrition: String(rows[i][2] || '').trim() };
    }
  }
  cache.put('MEAL_DATA', JSON.stringify(meals), 600);
  return meals;
}

function saveMealData(parsedMeals) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '다른 작업 처리 중입니다.' }; }
  try {
    const ss = setupSpreadsheet();
    const mealSheet = ss.getSheetByName('급식식단');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const existingMeals = fetchMealData(ss);

    Object.assign(existingMeals, parsedMeals);

    const rowsToSave = [['일자', '식단메뉴', '영양칼로리', '업데이트일시']];
    const allDates = Object.keys(existingMeals).sort();
    allDates.forEach(dKey => {
      rowsToSave.push([dKey, existingMeals[dKey].menu, existingMeals[dKey].nutrition, nowStr]);
    });

    mealSheet.clearContents();
    mealSheet.getRange(1, 1, rowsToSave.length, 4).setValues(rowsToSave);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().put('MEAL_DATA', JSON.stringify(existingMeals), 600);
    return { success: true, count: Object.keys(parsedMeals).length, allMeals: existingMeals };
  } finally {
    lock.releaseLock();
  }
}

function fetchNoticeAndDDayData(ssInstance) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('BOARD_DATA');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const ss = ssInstance || setupSpreadsheet();
  const noticeSheet = ss.getSheetByName('알림장');
  const result = { mainTitle: '오늘의 학급 알림장', ddays: [], notices: [] };
  if (!noticeSheet || noticeSheet.getLastRow() <= 1) return result;

  const rows = noticeSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const type = String(rows[i][0] || '').trim();
    const title = String(rows[i][1] || '').trim();
    const contentOrDate = rows[i][2];
    const tag = String(rows[i][3] || '').trim();

    if (type === 'TITLE') {
      result.mainTitle = String(contentOrDate || '오늘의 학급 알림장').trim();
    } else if (type === 'D_DAY') {
      const dateStr = (contentOrDate instanceof Date) 
        ? Utilities.formatDate(contentOrDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') 
        : String(contentOrDate).trim();
      if (title && dateStr) result.ddays.push({ id: i, title: title, targetDate: dateStr, tag: tag || '디데이' });
    } else if (type === 'NOTICE') {
      if (title || contentOrDate) result.notices.push({ id: i, title: title, content: String(contentOrDate || ''), tag: tag || '공지' });
    }
  }
  cache.put('BOARD_DATA', JSON.stringify(result), 600);
  return result;
}

function saveNoticeAndDDayData(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '다른 작업 처리 중입니다.' }; }
  try {
    const ss = setupSpreadsheet();
    const noticeSheet = ss.getSheetByName('알림장');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const rowsToSave = [['구분', '항목/제목', '날짜/내용', '추가정보', '업데이트일시']];
    const mainTitle = payload.mainTitle ? payload.mainTitle.trim() : '오늘의 학급 알림장';
    rowsToSave.push(['TITLE', '메인제목', mainTitle, '', nowStr]);

    if (Array.isArray(payload.ddays)) {
      payload.ddays.forEach(d => {
        if (d.title && d.targetDate) rowsToSave.push(['D_DAY', d.title.trim(), d.targetDate.trim(), (d.tag || '디데이').trim(), nowStr]);
      });
    }
    if (Array.isArray(payload.notices)) {
      payload.notices.forEach(n => {
        if (n.title || n.content) rowsToSave.push(['NOTICE', n.title.trim(), n.content.trim(), (n.tag || '일반').trim(), nowStr]);
      });
    }

    noticeSheet.clearContents();
    noticeSheet.getRange(1, 1, rowsToSave.length, 5).setValues(rowsToSave);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('BOARD_DATA');
    const updatedData = fetchNoticeAndDDayData(ss);
    return { success: true, message: '알림장 및 D-Day 설정이 저장되었습니다.', data: updatedData };
  } finally {
    lock.releaseLock();
  }
}

function fetchAllSpecialRecords(ssInstance) {
  try {
    const ss = ssInstance || setupSpreadsheet();
    const recordSheet = ss.getSheetByName('출결기록');
    const onlineReports = fetchOnlineReportsMap(ss);
    if (!recordSheet || recordSheet.getLastRow() <= 1) return [];

    const rows = recordSheet.getDataRange().getValues();
    const colMap = getColumnIndexMap(rows[0]);
    const specials = [];

    for (let i = 1; i < rows.length; i++) {
      const status = String(rows[i][colMap.status] || '출석').trim();
      if (status !== '출석' && status !== '미출석' && status !== '') {
        const rawDate = rows[i][colMap.date];
        const rDate = (rawDate instanceof Date) 
          ? Utilities.formatDate(rawDate, 'Asia/Seoul', 'yyyy-MM-dd') 
          : String(rawDate).trim();
        const sId = parseInt(rows[i][colMap.id], 10);
        if (!isNaN(sId)) {
          const onlineInfo = (onlineReports[sId] && onlineReports[sId][rDate]) ? onlineReports[sId][rDate] : null;
          specials.push({
            studentId: sId,
            date: rDate,
            time: formatDateToCustomString(rows[i][colMap.time]),
            status: status,
            period: (colMap.period !== -1 && rows[i][colMap.period]) ? String(rows[i][colMap.period]) : '종일',
            reason: (colMap.reason !== -1 && rows[i][colMap.reason]) ? String(rows[i][colMap.reason]) : '-',
            onlineReport: onlineInfo
          });
        }
      }
    }
    specials.sort((a, b) => b.date.localeCompare(a.date));
    return specials;
  } catch (err) {
    return [];
  }
}

function getStudentAttendanceHistory(studentId) {
  try {
    const ss = setupSpreadsheet();
    const recordSheet = ss.getSheetByName('출결기록');
    const onlineReports = fetchOnlineReportsMap(ss);
    if (!recordSheet || recordSheet.getLastRow() <= 1) return { success: true, specialRecords: [] };

    const sId = parseInt(studentId, 10);
    const rows = recordSheet.getDataRange().getValues();
    const colMap = getColumnIndexMap(rows[0]);
    const specials = [];

    for (let i = 1; i < rows.length; i++) {
      if (parseInt(rows[i][colMap.id], 10) === sId) {
        const status = String(rows[i][colMap.status] || '출석').trim();
        if (status !== '출석' && status !== '미출석' && status !== '') {
          const rawDate = rows[i][colMap.date];
          const rDate = (rawDate instanceof Date) 
            ? Utilities.formatDate(rawDate, 'Asia/Seoul', 'yyyy-MM-dd') 
            : String(rawDate).trim();

          const onlineInfo = (onlineReports[sId] && onlineReports[sId][rDate]) ? onlineReports[sId][rDate] : null;
          specials.push({
            studentId: sId,
            date: rDate,
            time: formatDateToCustomString(rows[i][colMap.time]),
            status: status,
            period: (colMap.period !== -1 && rows[i][colMap.period]) ? String(rows[i][colMap.period]) : '종일',
            reason: (colMap.reason !== -1 && rows[i][colMap.reason]) ? String(rows[i][colMap.reason]) : '-',
            onlineReport: onlineInfo
          });
        }
      }
    }
    specials.sort((a, b) => b.date.localeCompare(a.date));
    return { success: true, specialRecords: specials };
  } catch (err) {
    return { success: false, message: err.toString(), specialRecords: [] };
  }
}

function getInitialData(targetTodayStr) {
  const ss = setupSpreadsheet();
  const students = getStudentMasterData(ss);
  const config = {
    neisPeriod: '2026.03.02. - 2026.07.31.',
    schoolLat: 37.278698087555604,
    schoolLng: 127.45853065238362,
    radius: 250,
    dailyPin: '',
    startTime: '08:00',
    endTime: '08:50'
  };

  const configSheet = ss.getSheetByName('설정');
  if (configSheet && configSheet.getLastRow() > 1) {
    const configRows = configSheet.getDataRange().getValues();
    for (let k = 1; k < configRows.length; k++) {
      const key = String(configRows[k][0]).trim();
      const val = String(configRows[k][1]).trim();
      if (key === '나이스집계기간') config.neisPeriod = val;
      if (key === '오늘출석핀') config.dailyPin = val;
      if (key === '학교위도') config.schoolLat = parseFloat(val);
      if (key === '학교경도') config.schoolLng = parseFloat(val);
      if (key === '허용반경m') config.radius = parseInt(val, 10);
      if (key === '출석시작시간') config.startTime = val;
      if (key === '출석마감시간') config.endTime = val;
    }
  }

  const neisSheet = ss.getSheetByName('나이스누적');
  const neisData = {};
  if (neisSheet && neisSheet.getLastRow() > 1) {
    const neisRows = neisSheet.getDataRange().getValues();
    for (let m = 1; m < neisRows.length; m++) {
      const sNum = parseInt(neisRows[m][0], 10);
      if (!isNaN(sNum)) {
        neisData[sNum] = {
          name: String(neisRows[m][1] || ''),
          absent: { 질병: Number(neisRows[m][2]) || 0, 미인정: Number(neisRows[m][3]) || 0, 기타: Number(neisRows[m][4]) || 0, 인정: Number(neisRows[m][5]) || 0 },
          late: { 질병: Number(neisRows[m][6]) || 0, 미인정: Number(neisRows[m][7]) || 0, 기타: Number(neisRows[m][8]) || 0, 인정: Number(neisRows[m][9]) || 0 },
          early: { 질병: Number(neisRows[m][10]) || 0, 미인정: Number(neisRows[m][11]) || 0, 기타: Number(neisRows[m][12]) || 0, 인정: Number(neisRows[m][13]) || 0 },
          result: { 질병: Number(neisRows[m][14]) || 0, 미인정: Number(neisRows[m][15]) || 0, 기타: Number(neisRows[m][16]) || 0, 인정: Number(neisRows[m][17]) || 0 }
        };
      }
    }
  }

  const todayStr = targetTodayStr || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const currentYM = todayStr.substring(0, 7);
  const currentMonthMatrix = getMonthAttendanceMatrix(currentYM);

  return {
    success: true,
    students: students,
    config: config,
    neisData: neisData,
    onlineReports: fetchOnlineReportsMap(ss),
    todayRecords: fetchTodayAttendanceMap(todayStr, ss),
    allSpecialRecords: fetchAllSpecialRecords(ss),
    meals: fetchMealData(ss),
    boardData: fetchNoticeAndDDayData(ss),
    todayStr: todayStr,
    currentMonthMatrix: currentMonthMatrix ? currentMonthMatrix.records : {}
  };
}

function saveTimeLockSettings(startTime, endTime, radius) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '처리 중 오류가 발생했습니다.' }; }
  try {
    const ss = setupSpreadsheet();
    const configSheet = ss.getSheetByName('설정');
    const rows = configSheet.getDataRange().getValues();

    const updatedKeys = [];
    for (let i = 1; i < rows.length; i++) {
      const key = String(rows[i][0]).trim();
      if (key === '출석시작시간') { configSheet.getRange(i + 1, 2).setValue(startTime); updatedKeys.push(key); }
      if (key === '출석마감시간') { configSheet.getRange(i + 1, 2).setValue(endTime); updatedKeys.push(key); }
      if (key === '허용반경m') { configSheet.getRange(i + 1, 2).setValue(radius); updatedKeys.push(key); }
    }
    if (!updatedKeys.includes('출석시작시간')) configSheet.appendRow(['출석시작시간', startTime]);
    if (!updatedKeys.includes('출석마감시간')) configSheet.appendRow(['출석마감시간', endTime]);
    if (!updatedKeys.includes('허용반경m')) configSheet.appendRow(['허용반경m', radius]);
    SpreadsheetApp.flush();
    invalidateFastCache();
    return { success: true, message: '출석 시간 및 반경 설정이 저장되었습니다.' };
  } finally {
    lock.releaseLock();
  }
}

function getTeacherPassword(ssInstance) {
  try {
    const propPin = PropertiesService.getScriptProperties().getProperty('TEACHER_PIN');
    if (propPin && propPin.trim()) return propPin.trim();
  } catch (e) {}

  try {
    const ss = ssInstance || setupSpreadsheet();
    const configSheet = ss.getSheetByName('설정');
    if (configSheet && configSheet.getLastRow() > 1) {
      const rows = configSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === '교사비밀번호') {
          const val = String(rows[i][1]).trim();
          if (val) return val;
        }
      }
    }
  } catch (e) {}

  return '19650018';
}

function verifyTeacherPassword(inputPwd) {
  const targetPin = getTeacherPassword();
  return { success: (String(inputPwd).trim() === targetPin) };
}

function updateTeacherPassword(currentPassword, newPassword) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '처리 중 오류가 발생했습니다.' }; }
  try {
    const targetPin = getTeacherPassword();
    if (String(currentPassword).trim() !== targetPin) {
      return { success: false, message: '현재 교사 비밀번호가 일치하지 않습니다.' };
    }
    const cleanNewPwd = String(newPassword).trim();
    if (!cleanNewPwd || cleanNewPwd.length < 4) {
      return { success: false, message: '새 비밀번호는 4자리 이상이어야 합니다.' };
    }

    try {
      PropertiesService.getScriptProperties().setProperty('TEACHER_PIN', cleanNewPwd);
    } catch (e) {}

    const ss = setupSpreadsheet();
    const configSheet = ss.getSheetByName('설정');
    if (configSheet) {
      const rows = configSheet.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === '교사비밀번호') {
          configSheet.getRange(i + 1, 2).setValue(cleanNewPwd);
          found = true;
          break;
        }
      }
      if (!found) {
        configSheet.appendRow(['교사비밀번호', cleanNewPwd]);
      }
      SpreadsheetApp.flush();
    }
    return { success: true, message: '교사 비밀번호가 성공적으로 변경되었습니다.' };
  } finally {
    lock.releaseLock();
  }
}

function getMonthAttendanceMatrix(yearMonth) {
  const ss = setupSpreadsheet();
  const recordSheet = ss.getSheetByName('출결기록');
  const onlineReports = fetchOnlineReportsMap(ss);
  const monthRecords = {};

  if (!recordSheet || recordSheet.getLastRow() <= 1) return { success: true, yearMonth: yearMonth, records: {}, onlineReports: onlineReports };

  const rows = recordSheet.getDataRange().getValues();
  const colMap = getColumnIndexMap(rows[0]);

  for (let i = 1; i < rows.length; i++) {
    const rawDate = rows[i][colMap.date];
    const rDate = (rawDate instanceof Date) 
      ? Utilities.formatDate(rawDate, 'Asia/Seoul', 'yyyy-MM-dd') 
      : String(rawDate).trim();

    if (rDate.startsWith(yearMonth)) {
      const sId = parseInt(rows[i][colMap.id], 10);
      if (!isNaN(sId)) {
        if (!monthRecords[sId]) monthRecords[sId] = {};
        const onlineInfo = (onlineReports[sId] && onlineReports[sId][rDate]) ? onlineReports[sId][rDate] : null;

        monthRecords[sId][rDate] = {
          time: formatDateToCustomString(rows[i][colMap.time]),
          status: String(rows[i][colMap.status] || '출석'),
          period: (colMap.period !== -1 && rows[i][colMap.period]) ? String(rows[i][colMap.period]) : '종일',
          reason: (colMap.reason !== -1 && rows[i][colMap.reason]) ? String(rows[i][colMap.reason]) : '',
          gpsAuth: (colMap.auth !== -1 && rows[i][colMap.auth]) ? String(rows[i][colMap.auth]) : '',
          onlineReport: onlineInfo
        };
      }
    }
  }
  return { success: true, yearMonth: yearMonth, records: monthRecords, onlineReports: onlineReports };
}

function saveNeisData(period, parsedData) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '처리 중 오류가 발생했습니다.' }; }
  try {
    const ss = setupSpreadsheet();
    const neisSheet = ss.getSheetByName('나이스누적');
    const configSheet = ss.getSheetByName('설정');

    if (configSheet) {
      const configRows = configSheet.getDataRange().getValues();
      for (let i = 1; i < configRows.length; i++) {
        if (configRows[i][0] === '나이스집계기간') { 
          configSheet.getRange(i + 1, 2).setValue(period); 
          break; 
        }
      }
    }

    const students = getStudentMasterData(ss);
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const rowsToSave = [[
      '번호', '이름', 
      '결석_질병', '결석_미인정', '결석_기타', '결석_인정', 
      '지각_질병', '지각_미인정', '지각_기타', '지각_인정', 
      '조퇴_질병', '조퇴_미인정', '조퇴_기타', '조퇴_인정', 
      '결과_질병', '결과_미인정', '결과_기타', '결과_인정', 
      '업데이트일시', '집계기간'
    ]];

    students.forEach(st => {
      const d = parsedData[st.id] || parsedData[String(st.id)] || {};
      const ab = d.absent || { 질병: 0, 미인정: 0, 기타: 0, 인정: 0 };
      const la = d.late || { 질병: 0, 미인정: 0, 기타: 0, 인정: 0 };
      const ea = d.early || { 질병: 0, 미인정: 0, 기타: 0, 인정: 0 };
      const re = d.result || { 질병: 0, 미인정: 0, 기타: 0, 인정: 0 };

      rowsToSave.push([
        st.id, d.name || st.name,
        Number(ab.질병) || 0, Number(ab.미인정) || 0, Number(ab.기타) || 0, Number(ab.인정) || 0,
        Number(la.질병) || 0, Number(la.미인정) || 0, Number(la.기타) || 0, Number(la.인정) || 0,
        Number(ea.질병) || 0, Number(ea.미인정) || 0, Number(ea.기타) || 0, Number(ea.인정) || 0,
        Number(re.질병) || 0, Number(re.미인정) || 0, Number(re.기타) || 0, Number(re.인정) || 0,
        nowStr, period
      ]);
    });
    neisSheet.getRange(1, 1, rowsToSave.length, rowsToSave[0].length).setValues(rowsToSave);
    SpreadsheetApp.flush();
    return { success: true, period: period };
  } finally {
    lock.releaseLock();
  }
}

function updateStudentRoster(rosterList) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '처리 중 오류가 발생했습니다.' }; }
  try {
    const ss = setupSpreadsheet();
    const studentSheet = ss.getSheetByName('학생명단');
    const writeData = [['번호', '이름', '개인식별번호']];

    rosterList.sort((a, b) => parseInt(a.id, 10) - parseInt(b.id, 10));
    rosterList.forEach(st => {
      const p = normalizePin(st.pin) || '0000';
      writeData.push([st.id, String(st.name).trim(), "'" + p]);
    });

    studentSheet.clearContents();
    studentSheet.getRange(1, 1, writeData.length, 3).setValues(writeData);
    SpreadsheetApp.flush();
    invalidateFastCache();
    return { success: true, message: '학생 명단이 저장되었습니다.' };
  } finally {
    lock.releaseLock();
  }
}

function getTodayAttendanceStatus(targetTodayStr) {
  const todayStr = targetTodayStr || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  return { success: true, todayStr: todayStr, records: fetchTodayAttendanceMap(todayStr) };
}

function generateDailyPin() {
  const ss = setupSpreadsheet();
  const configSheet = ss.getSheetByName('설정');
  const pin = Math.floor(1000 + Math.random() * 9000).toString();
  const configRows = configSheet.getDataRange().getValues();

  for (let i = 1; i < configRows.length; i++) {
    if (String(configRows[i][0]).trim() === '오늘출석핀') {
      configSheet.getRange(i + 1, 2).setValue("'" + pin);
      SpreadsheetApp.flush();
      invalidateFastCache();
      return { success: true, pin: pin };
    }
  }
  configSheet.appendRow(['오늘출석핀', "'" + pin]);
  SpreadsheetApp.flush();
  invalidateFastCache();
  return { success: true, pin: pin };
}

function submitSelfAttendance(studentId, inputPin, userLat, userLng) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '잠시 후 다시 시도해주세요.' }; }
  try {
    const ss = setupSpreadsheet();
    const configSheet = ss.getSheetByName('설정');
    const configRows = configSheet.getDataRange().getValues();
    let targetPin = '', schoolLat = 37.278698087555604, schoolLng = 127.45853065238362, radius = 250;
    let startTime = '08:00', endTime = '08:50';

    for (let i = 1; i < configRows.length; i++) {
      const key = String(configRows[i][0]).trim();
      const val = String(configRows[i][1]).trim();
      if (key === '오늘출석핀') targetPin = val;
      if (key === '학교위도') schoolLat = parseFloat(val);
      if (key === '학교경도') schoolLng = parseFloat(val);
      if (key === '허용반경m') radius = parseInt(val, 10);
      if (key === '출석시작시간') startTime = val;
      if (key === '출석마감시간') endTime = val;
    }

    const now = new Date();
    // UTC offset to KST
    const kstNow = new Date(now.getTime() + (9 * 60 + now.getTimezoneOffset()) * 60000);
    const currentMinVal = kstNow.getHours() * 60 + kstNow.getMinutes();
    const sParts = startTime.split(':'), eParts = endTime.split(':');
    const startMinVal = parseInt(sParts[0], 10) * 60 + parseInt(sParts[1] || '0', 10);
    const endMinVal = parseInt(eParts[0], 10) * 60 + parseInt(eParts[1] || '0', 10);

    if (currentMinVal < startMinVal - 3 || currentMinVal > endMinVal + 3) {
      return { success: false, message: `현재는 자율 출석 시간이 아닙니다. (허용: ${startTime} ~ ${endTime})` };
    }

    const cleanTargetPin = normalizePin(targetPin);
    const cleanInputPin = normalizePin(inputPin);

    if (!cleanTargetPin) {
      return { success: false, message: '오늘의 출석 핀번호가 아직 생성되지 않았습니다. 담임선생님께 확인해주세요.' };
    }

    const students = getStudentMasterData(ss);
    const sId = parseInt(studentId, 10);
    const studentObj = students.find(s => s.id === sId);
    const studentName = studentObj ? studentObj.name : `학생${sId}`;
    const studentPin = (studentObj && studentObj.pin) ? normalizePin(studentObj.pin) : '0000';

    // 학생이 본인의 개인 식별번호(PIN 또는 기본 0000)를 입력했는지 확인하여 명확한 오류 안내 제공
    if ((studentPin === cleanInputPin || cleanInputPin === '0000') && cleanInputPin !== cleanTargetPin) {
      return { success: false, message: '개인 식별번호가 아닌, 칠판/교실 TV에 안내된 [오늘의 출석 핀번호 4자리]를 입력해주세요.' };
    }

    // 핀번호 일치 검사 (정확한 일치, 4자리 0패딩 호환, 정수값 호환)
    const isPinMatch = (cleanTargetPin === cleanInputPin) ||
                       (cleanTargetPin.padStart(4, '0') === cleanInputPin.padStart(4, '0')) ||
                       (parseInt(cleanTargetPin, 10) === parseInt(cleanInputPin, 10) && cleanInputPin.length >= 3);

    if (!isPinMatch) {
      return { success: false, message: '오늘의 출석 핀번호가 일치하지 않습니다. (칠판/화면의 4자리 번호를 확인하세요)' };
    }

    const distance = calculateDistance(userLat, userLng, schoolLat, schoolLng);
    if (distance > radius) {
      return { success: false, message: `학교 범위를 벗어났습니다. (${Math.round(distance)}m / 허용: ${radius}m)` };
    }

    const recordSheet = ss.getSheetByName('출결기록');
    const todayStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd');
    const nowStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const recordData = recordSheet.getDataRange().getValues();
    const colMap = getColumnIndexMap(recordData[0]);
    let foundRow = -1;

    for (let j = recordData.length - 1; j >= 1; j--) {
      const rawD = recordData[j][colMap.date];
      const rDate = (rawD instanceof Date) ? Utilities.formatDate(rawD, 'Asia/Seoul', 'yyyy-MM-dd') : String(rawD).trim();
      if (rDate === todayStr && parseInt(recordData[j][colMap.id], 10) === sId) { foundRow = j + 1; break; }
    }

    const authTxt = `인증됨(${Math.round(distance)}m)`;
    if (foundRow !== -1) {
      recordSheet.getRange(foundRow, colMap.time + 1).setValue(nowStr);
      recordSheet.getRange(foundRow, colMap.status + 1).setValue('출석');
      if (colMap.period !== -1) recordSheet.getRange(foundRow, colMap.period + 1).setValue('종일');
      if (colMap.reason !== -1) recordSheet.getRange(foundRow, colMap.reason + 1).setValue('자율 GPS 출석');
      if (colMap.auth !== -1) recordSheet.getRange(foundRow, colMap.auth + 1).setValue(authTxt);
    } else {
      recordSheet.appendRow([nowStr, todayStr, sId, studentName, '출석', '종일', '자율 GPS 출석', authTxt]);
    }
    SpreadsheetApp.flush();
    return { success: true, message: `출석 완료! (${startTime}~${endTime} 내 정상 처리)`, distance: Math.round(distance) };
  } finally {
    lock.releaseLock();
  }
}

function updateManualAttendance(studentId, status, period, reason, targetDate) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { success: false, message: '처리 중 오류가 발생했습니다.' }; }
  try {
    const ss = setupSpreadsheet();
    const students = getStudentMasterData(ss);
    const sId = parseInt(studentId, 10);
    const studentObj = students.find(s => s.id === sId);
    const studentName = studentObj ? studentObj.name : `학생${sId}`;

    const recordSheet = ss.getSheetByName('출결기록');
    const now = new Date();
    const effectiveDate = targetDate ? String(targetDate).trim() : Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd');
    const nowStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const recordData = recordSheet.getDataRange().getValues();
    const colMap = getColumnIndexMap(recordData[0]);
    let foundRow = -1;

    for (let j = recordData.length - 1; j >= 1; j--) {
      const rawD = recordData[j][colMap.date];
      const rDate = (rawD instanceof Date) ? Utilities.formatDate(rawD, 'Asia/Seoul', 'yyyy-MM-dd') : String(rawD).trim();
      if (rDate === effectiveDate && parseInt(recordData[j][colMap.id], 10) === sId) { foundRow = j + 1; break; }
    }

    if (status === '미출석' || status === '초기화') {
      if (foundRow !== -1) recordSheet.deleteRow(foundRow);
      SpreadsheetApp.flush();
      return { success: true, message: `${studentName} 학생의 [${effectiveDate}] 기록이 초기화되었습니다.` };
    }

    const finalPeriod = period || (status === '출석' ? '종일' : '조회');
    const finalReason = reason || '교사 직접 변경';
    if (foundRow !== -1) {
      recordSheet.getRange(foundRow, colMap.time + 1).setValue(nowStr);
      recordSheet.getRange(foundRow, colMap.status + 1).setValue(status);
      if (colMap.period !== -1) recordSheet.getRange(foundRow, colMap.period + 1).setValue(finalPeriod);
      if (colMap.reason !== -1) recordSheet.getRange(foundRow, colMap.reason + 1).setValue(finalReason);
      if (colMap.auth !== -1) recordSheet.getRange(foundRow, colMap.auth + 1).setValue('교사 수동 처리');
    } else {
      recordSheet.appendRow([nowStr, effectiveDate, sId, studentName, status, finalPeriod, finalReason, '교사 수동 처리']);
    }
    SpreadsheetApp.flush();
    return { success: true, message: `${studentName} 출결이 [${status}]로 변경되었습니다.` };
  } finally {
    lock.releaseLock();
  }
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371e3;
  const φ1 = lat1 * Math.PI / 180, φ2 = lat2 * Math.PI / 180;
  const Δφ = (lat2 - lat1) * Math.PI / 180, Δλ = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(Δφ / 2) * Math.sin(Δφ / 2) + Math.cos(φ1) * Math.cos(φ2) * Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
  return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}
