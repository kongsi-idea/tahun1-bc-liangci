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
    const validated = validateSubmission(payload);
    if (validated.error) return jsonOut(validated);
    return jsonOut(submitScore(validated));
  }
  // 'clear' 不接受公开的网页请求——任何人都能对着这个 Web App 网址发 POST，
  // 公开清空会让排行榜随时被恶意清空。要清空只能到 Apps Script 编辑器里手动
  // 执行 clearScores()，不透过这个入口。
  return jsonOut({ error: 'unknown action' });
}

// ------------------------------------------------------------------------
// 白名单验证：cls 只接受班级名单里的值，name 只接受对应班级名单里的真实姓名，
// score 限制在合理范围内的整数——避免任何人绕过前端直接对 API 发伪造成绩、
// 或用姓名字段塞入攻击字符串。跟 index.html 里的 CLASSES 保持同一份名单，
// 前端加/减学生时记得这边也要同步更新。
// ------------------------------------------------------------------------
const CLASSES = {
  '1I': ['谢佳宸','钟乐鹏','王梓宇','颜杰森','何承傧','李铭澔','许宸赫','许志安','顾宇乐','赖军暤','林安晟','罗茂洋','黄凯轩','郭瑞杰','董乐','黄明泽','吴昱杰','刘伊祎','陈偌宁','钟畇乐','吴辰娜','吴钫淣','吴钫嗪','马颖婕','许昕宁','李雨瞳','林柃希','王菱敏','刘乐蒽','张敏晏','张恩珣','黄湘霖','冯馨敏','谢瑜晨','张乐妍'],
  '1G': ['王煜勛','陈恩康','周宇轩','黄毅翔','杨伟皓','杨正宇','林锦喆','廖柏权','卢教炡','阿迪','玛丁','黄铭浚','谢宇恒','孙昊','陈有程','韦凯俊','尤辰瑞','余俊宏','徐语涵','黄欣怡','陳淑雯','谢淑惠','罗瑾棠','哈拿','张歆宁','何佳倩','罗楚恩','龚伊恩','黎馨雅','李嘉怡','陈乐妮','诺阿拉花','潘妤瑄','恺拉阿菲亚','施允熙','陈愉媗','叶禹彤'],
};
const MAX_SCORE = 300; // 10 题 x 10 分 + 旗杆最高 80 分奖励，留一点余裕；超过这个数一定是伪造

function validateSubmission(payload) {
  const cls = payload.cls;
  const name = payload.name;
  const score = Number(payload.score);
  if (!CLASSES[cls]) return { error: 'invalid class' };
  if (CLASSES[cls].indexOf(name) === -1) return { error: 'invalid name' };
  if (!Number.isFinite(score) || score < 0 || score > MAX_SCORE || !Number.isInteger(score)) {
    return { error: 'invalid score' };
  }
  return { cls: cls, name: name, score: score };
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
