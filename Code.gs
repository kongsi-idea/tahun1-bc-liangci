// ========================================================================
// 量词大冒险 —— 排行榜后端（纯 JSON API，不含任何页面）
// 网页本体部署在 Vercel/GitHub（见同资料夹 index.html），
// 这个 Apps Script 只负责把分数写进 / 读出 Google Sheet。
// 部署方式见同资料夹《部署步骤.md》。
// ========================================================================

const SHEET_NAME = '分数记录';

// ------------------------------------------------------------------------
// 入口：GET 查排行榜，POST 提交分数
// ------------------------------------------------------------------------
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  if (action === 'leaderboard') {
    return jsonOut(getLeaderboard());
  }
  return jsonOut({ error: 'unknown action' });
}

function doPost(e) {
  let payload = {};
  try { payload = JSON.parse(e.postData.contents); } catch (err) { return jsonOut({ error: 'bad payload' }); }

  if (payload.action === 'submit') {
    return jsonOut(submitScore(payload));
  }
  if (payload.action === 'clear') {
    return jsonOut({ cleared: clearScores() });
  }
  return jsonOut({ error: 'unknown action' });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------------------
// Sheet 工具函数
// ------------------------------------------------------------------------
function getSheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SHEET_ID');
  let ss = null;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('量词大冒险 - 分数记录');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['班级', '姓名', '最高分', '作答次数', '更新时间']);
  }
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
  return sheet;
}

// ------------------------------------------------------------------------
// 提交分数：只保留每位学生的历史最高分 + 累计作答次数
// ------------------------------------------------------------------------
function submitScore(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.cls && data[i][1] === payload.name) { rowIndex = i; break; }
    }
    const now = new Date();
    if (rowIndex === -1) {
      sheet.appendRow([payload.cls, payload.name, payload.score, 1, now]);
      return { best: payload.score, attempts: 1 };
    } else {
      const prevBest = Number(data[rowIndex][2]) || 0;
      const prevAttempts = Number(data[rowIndex][3]) || 0;
      const best = Math.max(prevBest, payload.score);
      const attempts = prevAttempts + 1;
      sheet.getRange(rowIndex + 1, 3).setValue(best);
      sheet.getRange(rowIndex + 1, 4).setValue(attempts);
      sheet.getRange(rowIndex + 1, 5).setValue(now);
      return { best: best, attempts: attempts };
    }
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------------------
// 读取排行榜
// ------------------------------------------------------------------------
function getLeaderboard() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1)
      .filter(r => r[1])
      .map(r => ({ cls: r[0], name: r[1], score: Number(r[2]) || 0, attempts: Number(r[3]) || 0 }));
    rows.sort((a, b) => b.score - a.score);
    return rows;
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------------------
// 清空本轮所有分数（活动结束或重新测试时，在编辑器手动执行）
// ------------------------------------------------------------------------
function clearScores() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------------------
// 维护用：如果表格资料错乱需要重新开始，在编辑器里手动执行这个函数一次
// ------------------------------------------------------------------------
function resetDatabase() {
  PropertiesService.getScriptProperties().deleteProperty('SHEET_ID');
  Logger.log('SHEET_ID 已清除，下次有人提交分数时会自动建立一份全新的表格。');
}
