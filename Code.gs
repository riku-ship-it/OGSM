// ====================================================
//  OGSM 策略看板系統 - Google Apps Script 後端（多職員版）
//
//  工作表結構：每位職員一個工作表，工作表名稱 = 職員姓名
//
//  工作表欄位對應（1列 = 1筆行動項目，目標/支線資訊重複填入）：
//    A(0)  編號       → Objective id
//    B(1)  目標標題   → Objective title
//    C(2)  支線編號   → Goal id
//    D(3)  支線名稱   → Goal name
//    E(4)  進度       → Goal progress
//    F(5)  顏色       → Goal color
//    G(6)  行動編號   → Action id
//    H(7)  策略名稱   → Action strategy_name
//    I(8)  行動項目   → Action action_name
//    J(9)  負責人     → Action assignee
//    K(10) 開始日期   → Action start_date
//    L(11) 截止日期   → Action due_date
//    M(12) 行動進度   → Action progress
//    N(13) 狀態       → Action status
//    O(14) 交通燈     → Goal traffic_light
//    P(15) 截止日     → Goal deadline
//
//  部署方式：
//    發布 → 部署為 Web 應用程式
//    執行身分：我（試算表擁有者）
//    存取權限：所有人（含匿名）
// ====================================================

var SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
var HEADER_ROW = ['編號','目標標題','支線編號','支線名稱','進度','顏色','行動編號','策略名稱','行動項目','負責人','開始日期','截止日期','行動進度','狀態','交通燈','截止日','策略狀態','成功定義'];
var STATS_HEADER_ROW = ['職員','ID','上線日期','系統平台','對象','項目說明','計分標準','分數'];

// ── 取得或建立統計工作表 ──
function getOrCreateStatsSheet(ss) {
  var sheet = ss.getSheetByName('Stats');
  if (sheet) return sheet;
  var newSheet = ss.insertSheet('Stats');
  newSheet.appendRow(STATS_HEADER_ROW);
  return newSheet;
}

// ── 取得或建立職員工作表 ──
function getSheetForStaff(ss, staffName) {
  var sheet = ss.getSheetByName(staffName);
  if (sheet) return sheet;

  // 遷移：若請求 Riku 且存在舊的「工作表1」，自動重新命名
  if (staffName === 'Riku') {
    var legacy = ss.getSheetByName('工作表1');
    if (legacy) {
      legacy.setName('Riku');
      return legacy;
    }
  }

  // 建立新工作表並加上標題列
  var newSheet = ss.insertSheet(staffName);
  newSheet.appendRow(HEADER_ROW);
  return newSheet;
}

// ====================================================
//  GET：
//    ?api=1&action=staff_list → 回傳所有職員名稱（工作表名稱）
//    ?api=1&staff=Riku        → 回傳 Riku 的 OGSM 資料
// ====================================================
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.api) {
    return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('OGSM 策略看板')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ---- 回傳職員清單 ----
    if (e.parameter.action === 'staff_list') {
      // 遷移：若有「工作表1」且沒有「Riku」，先重新命名
      var legacy = ss.getSheetByName('工作表1');
      if (legacy && !ss.getSheetByName('Riku')) {
        legacy.setName('Riku');
      }
      var names = ss.getSheets().map(function(s) { return s.getName(); });
      return ContentService
        .createTextOutput(JSON.stringify({ staff: names }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---- 回傳統計上線項目 ----
    if (e.parameter.action === 'get_stats') {
      var statsSheet = getOrCreateStatsSheet(ss);
      var statsData = statsSheet.getDataRange().getValues();
      var items = [];
      var filterStaff = e.parameter.staff || '';
      for (var i = 1; i < statsData.length; i++) {
        var row = statsData[i];
        if (!row[0]) continue;
        if (filterStaff && String(row[0]) !== filterStaff) continue;
        items.push({
          staff:       String(row[0] || ''),
          id:          String(row[1] || ''),
          launchDate:  String(row[2] || ''),
          platform:    String(row[3] || ''),
          target:      String(row[4] || ''),
          description: String(row[5] || ''),
          type:        String(row[6] || ''),
          score:       Number(row[7]) || 0
        });
      }
      return ContentService
        .createTextOutput(JSON.stringify({ items: items }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---- 回傳指定職員的 OGSM 資料 ----
    var staffName = e.parameter.staff || 'Riku';
    var sheet = getSheetForStaff(ss, staffName);
    var data = sheet.getDataRange().getValues();

    var objMap   = {};
    var goalMap  = {};
    var stratMap = {};
    var actions  = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && row[0] !== 0 && !row[6] && row[6] !== 0) continue;

      var objId     = String(row[0] || '');
      var goalId    = String(row[2] || '');
      var actId     = String(row[6] || '');
      var stratName = String(row[7] || '');

      if (objId && !objMap[objId]) {
        objMap[objId] = { id: objId, title: String(row[1] || '') };
      }

      if (goalId && !goalMap[goalId]) {
        goalMap[goalId] = {
          id:            goalId,
          objective_id:  objId,
          name:          String(row[3] || ''),
          progress:      Number(row[4])  || 0,
          color:         String(row[5]  || 'blue').toLowerCase().trim(),
          traffic_light: String(row[14] || ''),
          deadline:      formatDate(row[15])
        };
      }

      if (goalId && stratName) {
        var stratKey = goalId + '||' + stratName;
        if (!stratMap[stratKey]) {
          stratMap[stratKey] = { goal_id: goalId, name: stratName, status: String(row[16] || ''), success_def: String(row[17] || '') };
        } else {
          if (!stratMap[stratKey].status && row[16]) stratMap[stratKey].status = String(row[16]);
          if (!stratMap[stratKey].success_def && row[17]) stratMap[stratKey].success_def = String(row[17]);
        }
      }

      if (actId) {
        actions.push({
          id:            actId,
          goal_id:       goalId,
          strategy_name: stratName,
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
      strategies: Object.values(stratMap),
      actions:    actions
    });

    return ContentService
      .createTextOutput(payload)
      .setMimeType(ContentService.MimeType.JSON);

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
//  POST：依 body.staff 決定使用哪張工作表
// ====================================================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var result;

    // ---- add_stats_item：新增統計上線項目 ----
    if (body.type === 'add_stats_item') {
      var statsSheet = getOrCreateStatsSheet(ss);
      statsSheet.appendRow([
        String(body.staff       || ''),
        String(body.id          || ''),
        String(body.launchDate  || ''),
        String(body.platform    || ''),
        String(body.target      || ''),
        String(body.description || ''),
        String(body.type_name   || ''),
        Number(body.score)      || 0
      ]);
      result = JSON.stringify({ success: true, message: '新增成功' });

    // ---- update_stats_item：更新統計上線項目 ----
    } else if (body.type === 'update_stats_item') {
      var statsSheet = getOrCreateStatsSheet(ss);
      var statsData = statsSheet.getDataRange().getValues();
      var updated = false;
      for (var i = 1; i < statsData.length; i++) {
        if (String(statsData[i][1]) === String(body.id)) {
          var rowNum = i + 1;
          statsSheet.getRange(rowNum, 3).setValue(String(body.launchDate  || ''));
          statsSheet.getRange(rowNum, 4).setValue(String(body.platform    || ''));
          statsSheet.getRange(rowNum, 5).setValue(String(body.target      || ''));
          statsSheet.getRange(rowNum, 6).setValue(String(body.description || ''));
          statsSheet.getRange(rowNum, 7).setValue(String(body.type_name   || ''));
          statsSheet.getRange(rowNum, 8).setValue(Number(body.score)      || 0);
          updated = true;
          break;
        }
      }
      result = JSON.stringify({ success: updated, message: updated ? '更新成功' : '找不到項目：' + body.id });

    // ---- delete_stats_item：刪除統計上線項目 ----
    } else if (body.type === 'delete_stats_item') {
      var statsSheet = getOrCreateStatsSheet(ss);
      var statsData = statsSheet.getDataRange().getValues();
      var rowsToDelete = [];
      for (var i = 1; i < statsData.length; i++) {
        if (String(statsData[i][1]) === String(body.id)) rowsToDelete.push(i + 1);
      }
      for (var j = rowsToDelete.length - 1; j >= 0; j--) statsSheet.deleteRow(rowsToDelete[j]);
      result = JSON.stringify({ success: rowsToDelete.length > 0, message: rowsToDelete.length > 0 ? '刪除成功' : '找不到項目：' + body.id });

    // ---- add_staff：新增職員工作表 ----
    } else if (body.type === 'add_staff') {
      var staffName = String(body.staff_name || '').trim();
      if (!staffName) throw new Error('職員名稱不能為空');
      var existing = ss.getSheetByName(staffName);
      if (existing) {
        result = JSON.stringify({ success: false, message: '職員已存在：' + staffName });
      } else {
        var newSheet = ss.insertSheet(staffName);
        newSheet.appendRow(HEADER_ROW);
        result = JSON.stringify({ success: true, message: '新增成功' });
      }

    // ---- delete_staff：刪除職員工作表 ----
    } else if (body.type === 'delete_staff') {
      var staffName = String(body.staff_name || '').trim();
      if (!staffName) throw new Error('職員名稱不能為空');
      var targetSheet = ss.getSheetByName(staffName);
      if (!targetSheet) {
        result = JSON.stringify({ success: false, message: '找不到職員：' + staffName });
      } else if (ss.getSheets().length <= 1) {
        result = JSON.stringify({ success: false, message: '無法刪除最後一個工作表' });
      } else {
        ss.deleteSheet(targetSheet);
        result = JSON.stringify({ success: true, message: '刪除成功' });
      }

    // ---- ai_chat：透過後端代理呼叫 AI API ----
    } else if (body.type === 'ai_chat') {
      var props     = PropertiesService.getScriptProperties();
      var apiKey    = props.getProperty('API_KEY');
      var chatbotId = props.getProperty('CHATBOT_ID');
      if (!apiKey || !chatbotId) {
        result = JSON.stringify({ success: false, error: '請先在指令碼屬性設定 API_KEY、CHATBOT_ID' });
      } else {
        try {
          // 取得職員 OGSM 資料並組成 context
          var aiStaff  = String(body.staff || 'Riku');
          var ogsmRows = getSheetForStaff(ss, aiStaff).getDataRange().getValues();
          var ctxLines = ['# Current User Data', '當前職員：' + aiStaff, '', '## OGSM 資料'];
          var seenObj = {}, seenGoal = {}, seenStrat = {};
          for (var r = 1; r < ogsmRows.length; r++) {
            var row = ogsmRows[r];
            if (!row[0] && !row[6]) continue;
            var oKey = String(row[0] || '');
            var gKey = String(row[2] || '');
            var sKey = gKey + '||' + String(row[7] || '');
            if (oKey && !seenObj[oKey])     { seenObj[oKey]   = true; ctxLines.push('- 目標：'   + String(row[1] || '')); }
            if (gKey && !seenGoal[gKey])    { seenGoal[gKey]  = true; ctxLines.push('  - 支線：' + String(row[3] || '') + '（進度 ' + (row[4] || 0) + '%）'); }
            if (row[7] && !seenStrat[sKey]) { seenStrat[sKey] = true; ctxLines.push('    - 策略：' + String(row[7]) + (row[16] ? '（' + row[16] + '）' : '')); }
            if (row[6]) { ctxLines.push('      - 行動：' + String(row[8] || '') + '（負責人：' + String(row[9] || '') + '，截止：' + formatDate(row[11]) + '，狀態：' + String(row[13] || '') + '）'); }
          }
          var fullMessage = ctxLines.join('\n') + '\n\n---\n\n' + String(body.message || '');

          var apiUrl = 'https://api.maiagent.ai/api/v1/chatbots/' + chatbotId + '/completions/';
          var aiRes = UrlFetchApp.fetch(apiUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Authorization': 'Api-Key ' + apiKey
            },
            payload: JSON.stringify(Object.assign(
              { message: { content: fullMessage }, is_streaming: false },
              body.conversationId ? { conversation: body.conversationId } : {}
            )),
            muteHttpExceptions: true
          });
          var aiStatus = aiRes.getResponseCode();
          var aiData   = JSON.parse(aiRes.getContentText());
          if (aiStatus === 200 || aiStatus === 201) {
            var reply = aiData.content ||
                        (aiData.message && aiData.message.content) ||
                        aiData.answer || aiData.text || aiData.reply ||
                        JSON.stringify(aiData);
            var convId = aiData.conversationId || aiData.conversation || null;
            result = JSON.stringify({ success: true, reply: reply, conversationId: convId });
          } else {
            result = JSON.stringify({ success: false, error: 'AI API 回傳錯誤 ' + aiStatus + ': ' + aiRes.getContentText() });
          }
        } catch (aiErr) {
          result = JSON.stringify({ success: false, error: aiErr.message });
        }
      }

    } else {
      // 所有其他操作使用 body.staff 指定的工作表
      var staffName = String(body.staff || 'Riku');
      var sheet = getSheetForStaff(ss, staffName);
      var data  = sheet.getDataRange().getValues();

      // ---- rename_objective ----
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

      // ---- rename_goal ----
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

      // ---- rename_strategy ----
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

      // ---- rename_action ----
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

      // ---- update_goal_deadline ----
      } else if (body.type === 'update_goal_deadline') {
        var goalId   = String(body.goal_id);
        var deadline = String(body.deadline || '');
        var count    = 0;
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][2]) === goalId) {
            sheet.getRange(i + 1, 16).setValue(deadline);
            count++;
          }
        }
        result = JSON.stringify({ success: count > 0, message: count > 0 ? '更新成功' : '找不到目標：' + goalId });

      // ---- update_goal_traffic ----
      } else if (body.type === 'update_goal_traffic') {
        var goalId = String(body.goal_id);
        var light  = String(body.traffic_light || 'green');
        var count  = 0;
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][2]) === goalId) {
            sheet.getRange(i + 1, 15).setValue(light);
            count++;
          }
        }
        result = JSON.stringify({ success: count > 0, message: count > 0 ? '更新成功' : '找不到目標：' + goalId });

      // ---- delete_action ----
      } else if (body.type === 'delete_action') {
        var targetId = String(body.action_id);
        var rowsToDelete = [];
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][6]) === targetId) rowsToDelete.push(i + 1);
        }
        for (var j = rowsToDelete.length - 1; j >= 0; j--) sheet.deleteRow(rowsToDelete[j]);
        result = JSON.stringify({ success: rowsToDelete.length > 0, message: rowsToDelete.length > 0 ? '刪除成功' : '找不到行動：' + targetId });

      // ---- delete_strategy ----
      } else if (body.type === 'delete_strategy') {
        var goalId    = String(body.goal_id);
        var stratName = String(body.strategy_name || '');
        var rowsToDelete = [];
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][2]) === goalId && String(data[i][7]) === stratName) rowsToDelete.push(i + 1);
        }
        for (var j = rowsToDelete.length - 1; j >= 0; j--) sheet.deleteRow(rowsToDelete[j]);
        result = JSON.stringify({ success: rowsToDelete.length > 0, message: rowsToDelete.length > 0 ? '刪除成功' : '找不到策略：' + stratName });

      // ---- delete_goal ----
      } else if (body.type === 'delete_goal') {
        var goalId = String(body.goal_id);
        var rowsToDelete = [];
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][2]) === goalId) rowsToDelete.push(i + 1);
        }
        for (var j = rowsToDelete.length - 1; j >= 0; j--) sheet.deleteRow(rowsToDelete[j]);
        result = JSON.stringify({ success: rowsToDelete.length > 0, message: rowsToDelete.length > 0 ? '刪除成功' : '找不到目標：' + goalId });

      // ---- add_goal ----
      } else if (body.type === 'add_goal') {
        var newRow = [
          String(body.obj_id    || ''),
          String(body.obj_title || ''),
          String(body.goal_id   || ''),
          String(body.goal_name || ''),
          Number(body.goal_progress) || 0,
          String(body.goal_color || 'blue'),
          '', '', '', '', '', '', 0, '', '',
          String(body.goal_deadline || ''),
          '', ''
        ];
        sheet.appendRow(newRow);
        result = JSON.stringify({ success: true, message: '新增成功' });

      // ---- add_action ----
      } else if (body.type === 'add_action') {
        var newRow = [
          String(body.obj_id        || ''),
          String(body.obj_title     || ''),
          String(body.goal_id       || ''),
          String(body.goal_name     || ''),
          Number(body.goal_progress) || 0,
          String(body.goal_color    || 'blue'),
          String(body.action_id     || ''),
          String(body.strategy_name || ''),
          String(body.action_name   || ''),
          String(body.assignee      || ''),
          '',
          String(body.due_date      || ''),
          Number(body.progress)     || 0,
          String(body.status        || '未開始'),
          '', '',
          '', ''
        ];
        sheet.appendRow(newRow);
        result = JSON.stringify({ success: true, message: '新增成功' });

      // ---- update_action ----
      } else if (body.type === 'update_action') {
        var targetId = String(body.id);
        var updated  = false;
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][6]) === targetId) {
            var rowNum = i + 1;
            if (body.strategy_name !== undefined) sheet.getRange(rowNum, 8).setValue(body.strategy_name);
            if (body.action_name !== undefined) sheet.getRange(rowNum, 9).setValue(body.action_name);
            if (body.assignee    !== undefined) sheet.getRange(rowNum, 10).setValue(body.assignee);
            if (body.due_date    !== undefined) sheet.getRange(rowNum, 12).setValue(body.due_date);
            if (body.progress    !== undefined) sheet.getRange(rowNum, 13).setValue(Number(body.progress));
            if (body.status      !== undefined) sheet.getRange(rowNum, 14).setValue(body.status);
            if (body.success_def !== undefined) sheet.getRange(rowNum, 18).setValue(body.success_def);
            updated = true;
            break;
          }
        }
        result = JSON.stringify({ success: updated, message: updated ? '更新成功' : '找不到行動編號：' + targetId });

      // ---- update_goal_color ----
      } else if (body.type === 'update_goal_color') {
        var goalId = String(body.goal_id);
        var color  = String(body.color || 'blue');
        var count  = 0;
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][2]) === goalId) {
            sheet.getRange(i + 1, 6).setValue(color);
            count++;
          }
        }
        result = JSON.stringify({ success: count > 0, message: count > 0 ? '更新成功' : '找不到目標：' + goalId });

      // ---- update_strategy_status ----
      } else if (body.type === 'update_strategy_status') {
        var goalId    = String(body.goal_id);
        var stratName = String(body.strategy_name || '');
        var status    = String(body.status || '進行中');
        var count     = 0;
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][2]) === goalId && String(data[i][7]) === stratName) {
            sheet.getRange(i + 1, 17).setValue(status);
            count++;
          }
        }
        result = JSON.stringify({ success: count > 0, message: count > 0 ? '更新成功' : '找不到策略' });

      // ---- update_strategy_success_def ----
      } else if (body.type === 'update_strategy_success_def') {
        var goalId     = String(body.goal_id);
        var stratName  = String(body.strategy_name || '');
        var successDef = String(body.success_def || '');
        var count      = 0;
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][2]) === goalId && String(data[i][7]) === stratName) {
            sheet.getRange(i + 1, 18).setValue(successDef);
            count++;
          }
        }
        result = JSON.stringify({ success: count > 0, message: count > 0 ? '更新成功' : '找不到策略' });

      // ---- reorder_goals ----
      } else if (body.type === 'reorder_goals') {
        var goalIds = (body.goal_ids || []).map(String);
        var rows = data.slice(1);
        var goalRowsMap = {};
        var seenGoals = [];
        rows.forEach(function(row) {
          var gid = String(row[2] || '');
          if (!goalRowsMap[gid]) { goalRowsMap[gid] = []; seenGoals.push(gid); }
          goalRowsMap[gid].push(row);
        });
        var reordered = [];
        goalIds.forEach(function(gid) {
          if (goalRowsMap[gid]) goalRowsMap[gid].forEach(function(r) { reordered.push(r); });
        });
        seenGoals.forEach(function(gid) {
          if (goalIds.indexOf(gid) === -1 && goalRowsMap[gid]) {
            goalRowsMap[gid].forEach(function(r) { reordered.push(r); });
          }
        });
        var numCols = HEADER_ROW.length;
        var numRows = sheet.getLastRow() - 1;
        if (numRows > 0) sheet.getRange(2, 1, numRows, numCols).clearContent();
        if (reordered.length > 0) sheet.getRange(2, 1, reordered.length, numCols).setValues(reordered);
        result = JSON.stringify({ success: true });

      // ---- reorder_strategies ----
      } else if (body.type === 'reorder_strategies') {
        var goalId = String(body.goal_id);
        var stratNames = (body.strategy_names || []).map(String);
        var rows = data.slice(1);
        var goalRowIndices = [];
        var stratRowsMap = {};
        var seenStrats = [];
        rows.forEach(function(row, i) {
          if (String(row[2]) === goalId) {
            goalRowIndices.push(i);
            var sn = String(row[7] || '');
            if (!stratRowsMap[sn]) { stratRowsMap[sn] = []; seenStrats.push(sn); }
            stratRowsMap[sn].push(row);
          }
        });
        var reorderedGoalRows = [];
        stratNames.forEach(function(sn) {
          if (stratRowsMap[sn]) stratRowsMap[sn].forEach(function(r) { reorderedGoalRows.push(r); });
        });
        seenStrats.forEach(function(sn) {
          if (stratNames.indexOf(sn) === -1 && stratRowsMap[sn]) {
            stratRowsMap[sn].forEach(function(r) { reorderedGoalRows.push(r); });
          }
        });
        var newRows = rows.slice();
        goalRowIndices.forEach(function(idx, i) {
          if (reorderedGoalRows[i]) newRows[idx] = reorderedGoalRows[i];
        });
        var numCols = HEADER_ROW.length;
        var numRows = sheet.getLastRow() - 1;
        if (numRows > 0) sheet.getRange(2, 1, numRows, numCols).clearContent();
        if (newRows.length > 0) sheet.getRange(2, 1, newRows.length, numCols).setValues(newRows);
        result = JSON.stringify({ success: true });

      // ---- reorder_actions ----
      } else if (body.type === 'reorder_actions') {
        var actionIds = (body.action_ids || []).map(String);
        var rows = data.slice(1);
        var actionRowIndices = [];
        var idToRow = {};
        rows.forEach(function(row, i) {
          var aid = String(row[6] || '');
          if (actionIds.indexOf(aid) !== -1) {
            actionRowIndices.push(i);
            idToRow[aid] = row;
          }
        });
        var reorderedActionRows = actionIds.map(function(aid) { return idToRow[aid]; }).filter(Boolean);
        var newRows = rows.slice();
        actionRowIndices.forEach(function(idx, i) {
          if (reorderedActionRows[i]) newRows[idx] = reorderedActionRows[i];
        });
        var numCols = HEADER_ROW.length;
        var numRows = sheet.getLastRow() - 1;
        if (numRows > 0) sheet.getRange(2, 1, numRows, numCols).clearContent();
        if (newRows.length > 0) sheet.getRange(2, 1, newRows.length, numCols).setValues(newRows);
        result = JSON.stringify({ success: true });

      // ---- 預設：依行動編號更新欄位 ----
      } else {
        var targetId = String(body.id);
        var updated  = false;
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][6]) === targetId) {
            var rowNum = i + 1;
            if (body.assignee !== undefined) sheet.getRange(rowNum, 10).setValue(body.assignee);
            if (body.due_date !== undefined) sheet.getRange(rowNum, 12).setValue(body.due_date);
            if (body.progress !== undefined) sheet.getRange(rowNum, 13).setValue(Number(body.progress));
            if (body.status   !== undefined) sheet.getRange(rowNum, 14).setValue(body.status);
            updated = true;
            break;
          }
        }
        result = JSON.stringify({ success: updated, message: updated ? '更新成功' : '找不到行動編號：' + targetId });
      }
    }

    return ContentService
      .createTextOutput(result)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ====================================================
//  HTML include 輔助函數
// ====================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====================================================
//  授權觸發函式（手動執行一次即可，之後可刪除）
//  用途：強制觸發 UrlFetchApp 的授權視窗
// ====================================================
function authorizeUrlFetch() {
  UrlFetchApp.fetch('https://www.google.com');
  Logger.log('UrlFetchApp 授權成功');
}
