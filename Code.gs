// ====================================================
//  OGSM 策略看板系統 - Google Apps Script 後端（單一工作表版）
//
//  工作表欄位對應（1列 = 1筆行動項目，目標/支線資訊重複填入）：
//    A(0)  編號       → Objective id
//    B(1)  目標標題   → Objective title
//    C(2)  支線編號   → Goal id
//    D(3)  支線名稱   → Goal name
//    E(4)  進度       → Goal progress
//    F(5)  顏色       → Goal color
//    G(6)  行動編號   → Action id  ← POST 以此欄定位列
//    H(7)  策略名稱   → Action strategy_name
//    I(8)  行動項目   → Action action_name
//    J(9)  負責人     → Action assignee      ← POST 更新
//    K(10) 開始日期   → Action start_date
//    L(11) 截止日期   → Action due_date      ← POST 更新
//    M(12) 行動進度   → Action progress      ← POST 更新
//    N(13) 狀態       → Action status        ← POST 更新
//
//  部署方式：
//    發布 → 部署為 Web 應用程式
//    執行身分：我（試算表擁有者）
//    存取權限：所有人（含匿名）
// ====================================================

// ★ 如果你的工作表名稱不是「工作表1」，請修改這裡
var SHEET_NAME = '工作表1';

// ====================================================
//  HTML include 輔助函數
//  用法：在 HTML 中以 <?!= include('css'); ?> 引入外部 HTML 檔案
// ====================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

var SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();


// ====================================================
//  GET：讀取單一工作表，回傳 { objectives, goals, actions }
// ====================================================
function doGet(e) {
  if (e && e.parameter && !e.parameter.api) {
    return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('OGSM 策略看板')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('找不到工作表：' + SHEET_NAME);

    var data = sheet.getDataRange().getValues();

    var objMap    = {};  // key = Objective id
    var goalMap   = {};  // key = Goal id
    var actions   = [];

    for (var i = 1; i < data.length; i++) {   // 跳過第 0 列（標題）
      var row = data[i];

      // 跳過完全空白的列
      if (!row[0] && row[0] !== 0 && !row[6] && row[6] !== 0) continue;

      var objId  = String(row[0] || '');
      var goalId = String(row[2] || '');
      var actId  = String(row[6] || '');

      // ---- Objective 去重 ----
      if (objId && !objMap[objId]) {
        objMap[objId] = {
          id:    objId,
          title: String(row[1] || '')
        };
      }

      // ---- Goal 去重 ----
      if (goalId && !goalMap[goalId]) {
        goalMap[goalId] = {
          id:           goalId,
          objective_id: objId,
          name:         String(row[3] || ''),
          progress:     Number(row[4])  || 0,
          color:        String(row[5]  || 'blue').toLowerCase().trim()
        };
      }

      // ---- Action（每列都加） ----
      if (actId) {
        actions.push({
          id:            actId,
          goal_id:       goalId,
          strategy_name: String(row[7]  || ''),
          action_name:   String(row[8]  || ''),
          assignee:      String(row[9]  || ''),
          start_date:    formatDate(row[10]),
          due_date:      formatDate(row[11]),
          progress:      Number(row[12]) || 0,
          status:        String(row[13] || '未開始')
        });
      }
    }

    var payload = JSON.stringify({
      objectives: Object.values(objMap),
      goals:      Object.values(goalMap),
      actions:    actions
    });

    var output = ContentService
      .createTextOutput(payload)
      .setMimeType(ContentService.MimeType.JSON);

    return output;

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- 日期格式化 ----------
function formatDate(value) {
  if (!value) return '';
  if (value instanceof Date) {
    var y = value.getFullYear();
    var m = String(value.getMonth() + 1).padStart(2, '0');
    var d = String(value.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(value);
}

// ====================================================
//  POST：依「行動編號（G欄）」定位列，更新指定欄位
//  接收 JSON：{ id, progress, status, assignee, due_date }
//
//  欄位對應（1-indexed 給 getRange 使用）：
//    J = 欄 10 → 負責人
//    L = 欄 12 → 截止日期
//    M = 欄 13 → 行動進度
//    N = 欄 14 → 狀態
// ====================================================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('找不到工作表：' + SHEET_NAME);

    var data = sheet.getDataRange().getValues();
    var result;

    // ---- rename_objective：更新所有符合 obj_id 的列的 B 欄（目標標題）----
    if (body.type === 'rename_objective') {
      var objId    = String(body.obj_id);
      var newTitle = String(body.new_title || '').trim();
      var count    = 0;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === objId) {
          sheet.getRange(i + 1, 2).setValue(newTitle);
          count++;
        }
      }
      result = JSON.stringify({ success: count > 0, message: count > 0 ? '更新成功' : '找不到目標：' + objId });

    // ---- rename_goal：更新所有符合 goal_id 的列的 D 欄（支線名稱）----
    } else if (body.type === 'rename_goal') {
      var goalId  = String(body.goal_id);
      var newName = String(body.new_name || '').trim();
      var count   = 0;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][2]) === goalId) {
          sheet.getRange(i + 1, 4).setValue(newName);
          count++;
        }
      }
      result = JSON.stringify({ success: count > 0, message: count > 0 ? '更新成功' : '找不到支線：' + goalId });

    // ---- rename_strategy：更新符合 goal_id + old_name 的列的 H 欄（策略名稱）----
    } else if (body.type === 'rename_strategy') {
      var goalId  = String(body.goal_id);
      var oldName = String(body.old_name || '');
      var newName = String(body.new_name || '').trim();
      var count   = 0;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][2]) === goalId && String(data[i][7]) === oldName) {
          sheet.getRange(i + 1, 8).setValue(newName);
          count++;
        }
      }
      result = JSON.stringify({ success: count > 0, message: count > 0 ? '更新成功' : '找不到策略：' + oldName });

    // ---- rename_action：更新符合 action_id 的列的 I 欄（行動名稱）----
    } else if (body.type === 'rename_action') {
      var targetId = String(body.action_id);
      var newName  = String(body.new_name || '').trim();
      var updated  = false;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][6]) === targetId) {
          sheet.getRange(i + 1, 9).setValue(newName);
          updated = true;
          break;
        }
      }
      result = JSON.stringify({ success: updated, message: updated ? '更新成功' : '找不到行動：' + targetId });

    // ---- 預設：依行動編號更新行動欄位 ----
    } else {
      var targetId = String(body.id);
      var updated  = false;

      for (var i = 1; i < data.length; i++) {
        var actionId = String(data[i][6]); // G 欄（index 6）= 行動編號
        if (actionId === targetId) {
          var rowNum = i + 1;

          // J（欄 10）= 負責人
          if (body.assignee !== undefined) {
            sheet.getRange(rowNum, 10).setValue(body.assignee);
          }
          // L（欄 12）= 截止日期
          if (body.due_date !== undefined) {
            sheet.getRange(rowNum, 12).setValue(body.due_date);
          }
          // M（欄 13）= 行動進度
          if (body.progress !== undefined) {
            sheet.getRange(rowNum, 13).setValue(Number(body.progress));
          }
          // N（欄 14）= 狀態
          if (body.status !== undefined) {
            sheet.getRange(rowNum, 14).setValue(body.status);
          }

          updated = true;
          break;
        }
      }

      result = JSON.stringify({
        success: updated,
        message: updated ? '更新成功' : '找不到行動編號：' + targetId
      });
    }

    var output = ContentService
      .createTextOutput(result)
      .setMimeType(ContentService.MimeType.JSON);

    return output;

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
