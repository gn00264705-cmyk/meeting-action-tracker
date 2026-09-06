/**
 * 個人會議行事曆與改善事項追蹤系統
 * 綁定於試算表時會優先使用目前試算表；獨立專案可設定下列常數或指令碼屬性 SPREADSHEET_ID。
 */
const SPREADSHEET_ID = '';
const APP = {
  spreadsheetIdProperty: 'SPREADSHEET_ID',
  timezone: Session.getScriptTimeZone() || 'Asia/Taipei',
  sheets: {
    Meetings: ['meeting_id', 'date', 'weekday', 'start_time', 'end_time', 'title', 'created_at'],
    ActionItems: ['item_id', 'meeting_id', 'content', 'source_person', 'status', 'notes']
  },
  statuses: ['待處理', '處理中', '已完成']
};

/** 載入 Web 應用程式。 */
function doGet() {
  initializeDatabase_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('會議行事曆與改善事項')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 取得會議及其改善事項。
 * filters 可傳 { month: 'YYYY-MM' } 或 { startDate: 'YYYY-MM-DD', endDate: 'YYYY-MM-DD' }。
 */
function apiGetMeetingsWithItems(filters) {
  filters = filters || {};
  const db = initializeDatabase_();
  const meetings = readObjects_(db.meetings, APP.sheets.Meetings)
    .filter(function (meeting) { return isWithinRange_(meeting.date, filters); })
    .map(function (meeting) {
      meeting.items = [];
      return meeting;
    });
  const byMeetingId = meetings.reduce(function (result, meeting) {
    result[meeting.meeting_id] = meeting;
    return result;
  }, {});

  readObjects_(db.actionItems, APP.sheets.ActionItems).forEach(function (item) {
    if (byMeetingId[item.meeting_id]) byMeetingId[item.meeting_id].items.push(item);
  });

  meetings.forEach(function (meeting) {
    meeting.items.sort(function (a, b) { return String(a.item_id).localeCompare(String(b.item_id)); });
  });
  meetings.sort(function (a, b) {
    return (b.date + ' ' + b.start_time).localeCompare(a.date + ' ' + a.start_time);
  });
  return { meetings: meetings, statuses: APP.statuses };
}

/**
 * 取得跨月份的全部未結案改善事項。
 * 回傳的每一筆事項均附帶 meetingDate、meetingTitle 與 meetingStartTime，供全域管考介面直接使用。
 */
function apiGetAllPendingActionItems() {
  const db = initializeDatabase_();
  const meetingsById = readObjects_(db.meetings, APP.sheets.Meetings).reduce(function (index, meeting) {
    index[meeting.meeting_id] = meeting;
    return index;
  }, {});

  return readObjects_(db.actionItems, APP.sheets.ActionItems)
    .filter(function (item) { return item.status !== '已完成'; })
    .map(function (item) {
      const meeting = meetingsById[item.meeting_id];
      return Object.assign({}, item, {
        meetingDate: meeting ? meeting.date : '',
        meetingTitle: meeting ? meeting.title : '（找不到所屬會議）',
        meetingStartTime: meeting ? meeting.start_time : ''
      });
    })
    .sort(function (a, b) {
      return (b.meetingDate + ' ' + b.meetingStartTime).localeCompare(a.meetingDate + ' ' + a.meetingStartTime);
    });
}

/**
 * 取得跨月份的全部已完成事項；以會議日期與會議名稱補足來源資訊。
 * 完成時間未存於既有資料表，因此時間篩選以 meetingDate（原始開會日期）為準。
 */
function apiGetAllCompletedActionItems() {
  const db = initializeDatabase_();
  const meetingsById = readObjects_(db.meetings, APP.sheets.Meetings).reduce(function (index, meeting) {
    index[meeting.meeting_id] = meeting;
    return index;
  }, {});

  return readObjects_(db.actionItems, APP.sheets.ActionItems)
    .filter(function (item) { return item.status === '已完成'; })
    .map(function (item) {
      const meeting = meetingsById[item.meeting_id];
      return Object.assign({}, item, {
        meetingDate: meeting ? meeting.date : '',
        meetingTitle: meeting ? meeting.title : '（找不到所屬會議）',
        meetingStartTime: meeting ? meeting.start_time : ''
      });
    })
    .sort(function (a, b) {
      return (b.meetingDate + ' ' + b.meetingStartTime).localeCompare(a.meetingDate + ' ' + a.meetingStartTime);
    });
}

/** 新增或更新會議主檔；更新時請傳 meeting_id。 */
function apiSaveMeeting(meetingData) {
  meetingData = parsePayload_(meetingData, '會議資料');
  const date = normalizeDate_(meetingData.date);
  const startTime = normalizeTime_(meetingData.start_time, '開始時間');
  const endTime = normalizeTime_(meetingData.end_time, '結束時間');
  const title = requiredText_(meetingData.title, '會議名稱');
  if (endTime <= startTime) throw new Error('結束時間必須晚於開始時間。');

  // 寫入前明確確認 Meetings 存在；首次執行時自動建立兩個分頁與表頭。
  const spreadsheet = getSpreadsheet_();
  if (!spreadsheet.getSheetByName('Meetings')) initSheets();
  const db = initializeDatabase_();
  const existing = meetingData.meeting_id ? findRowByValue_(db.meetings, 1, meetingData.meeting_id) : 0;
  if (meetingData.meeting_id && !existing) throw new Error('找不到欲更新的會議。');

  const id = existing ? String(meetingData.meeting_id) : Utilities.getUuid();
  const createdAt = existing
    ? db.meetings.getRange(existing, 7).getDisplayValue()
    : Utilities.formatDate(new Date(), APP.timezone, "yyyy-MM-dd'T'HH:mm:ss");
  const row = [id, date, getChineseWeekday_(date), startTime, endTime, title, createdAt];
  const targetRow = existing || db.meetings.getLastRow() + 1;
  // 日期與時間以文字儲存，避免試算表地區格式改變 API 所需的 YYYY-MM-DD／HH:mm。
  db.meetings.getRange(targetRow, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  const persistedId = db.meetings.getRange(targetRow, 1).getDisplayValue();
  if (persistedId !== id) throw new Error('會議資料寫入驗證失敗，請檢查試算表權限與分頁設定。');
  return { meeting_id: id, message: existing ? '會議已更新。' : '會議已新增。' };
}

/** 新增或更新指定會議底下的改善事項；更新時請傳 item_id。 */
function apiSaveActionItem(itemData) {
  itemData = parsePayload_(itemData, '改善事項資料');
  const meetingId = requiredText_(itemData.meeting_id, '所屬會議');
  const content = requiredText_(itemData.content, '改善事項內容');
  const sourcePerson = String(itemData.source_person || '').trim();
  const status = itemData.status || '待處理';
  const notes = String(itemData.notes || '').trim();
  if (APP.statuses.indexOf(status) === -1) throw new Error('改善事項狀態不正確。');

  const db = initializeDatabase_();
  if (!findRowByValue_(db.meetings, 1, meetingId)) throw new Error('所屬會議不存在。');
  const existing = itemData.item_id ? findRowByValue_(db.actionItems, 1, itemData.item_id) : 0;
  if (itemData.item_id && !existing) throw new Error('找不到欲更新的改善事項。');

  const id = existing ? String(itemData.item_id) : Utilities.getUuid();
  const row = [id, meetingId, content, sourcePerson, status, notes];
  const targetRow = existing || db.actionItems.getLastRow() + 1;
  db.actionItems.getRange(targetRow, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  return { item_id: id, message: existing ? '改善事項已更新。' : '改善事項已新增。' };
}

/**
 * 批次匯入會議與其改善事項。前端必須傳入已正規化的會議陣列，
 * 每筆格式為 { date, start_time, end_time, title, items: [{ content, source_person, status, notes }] }。
 * 此函式刻意只使用兩次 setValues()，避免在迴圈內逐筆 appendRow() 而造成效能瓶頸。
 */
function apiBatchImportMeetings(payloadList) {
  const meetingsToImport = parseArrayPayload_(payloadList, '批次匯入資料');
  if (!meetingsToImport.length) throw new Error('沒有可匯入的會議資料。');

  const db = initializeDatabase_();
  const createdAt = Utilities.formatDate(new Date(), APP.timezone, "yyyy-MM-dd'T'HH:mm:ss");
  const meetingRows = [];
  const itemRows = [];

  // 先完成所有資料驗證與記憶體整理，任一筆不合格時不寫入任何資料。
  meetingsToImport.forEach(function (meetingData, meetingIndex) {
    if (!meetingData || typeof meetingData !== 'object' || Array.isArray(meetingData)) {
      throw new Error('第 ' + (meetingIndex + 1) + ' 筆會議資料格式不正確。');
    }
    const date = normalizeDate_(meetingData.date);
    const startTime = normalizeTime_(meetingData.start_time, '第 ' + (meetingIndex + 1) + ' 筆開始時間');
    const endTime = normalizeTime_(meetingData.end_time, '第 ' + (meetingIndex + 1) + ' 筆結束時間');
    const title = requiredText_(meetingData.title, '第 ' + (meetingIndex + 1) + ' 筆會議名稱');
    if (endTime <= startTime) throw new Error('第 ' + (meetingIndex + 1) + ' 筆會議的結束時間必須晚於開始時間。');

    const meetingId = Utilities.getUuid();
    meetingRows.push([meetingId, date, getChineseWeekday_(date), startTime, endTime, title, createdAt]);

    if (meetingData.items !== undefined && !Array.isArray(meetingData.items)) {
      throw new Error('第 ' + (meetingIndex + 1) + ' 筆會議的改善事項必須是陣列。');
    }
    (meetingData.items || []).forEach(function (itemData, itemIndex) {
      if (!itemData || typeof itemData !== 'object' || Array.isArray(itemData)) {
        throw new Error('第 ' + (meetingIndex + 1) + ' 筆會議的第 ' + (itemIndex + 1) + ' 筆改善事項格式不正確。');
      }
      const content = requiredText_(itemData.content, '第 ' + (meetingIndex + 1) + ' 筆會議的第 ' + (itemIndex + 1) + ' 筆改善事項內容');
      const status = String(itemData.status || '待處理').trim();
      if (APP.statuses.indexOf(status) === -1) throw new Error('第 ' + (meetingIndex + 1) + ' 筆會議的改善事項狀態不正確。');
      itemRows.push([
        Utilities.getUuid(),
        meetingId,
        content,
        String(itemData.source_person || '').trim(),
        status,
        String(itemData.notes || '').trim()
      ]);
    });
  });

  const meetingStartRow = db.meetings.getLastRow() + 1;
  const itemStartRow = db.actionItems.getLastRow() + 1;
  db.meetings.getRange(meetingStartRow, 1, meetingRows.length, APP.sheets.Meetings.length).setNumberFormat('@').setValues(meetingRows);
  if (itemRows.length) {
    db.actionItems.getRange(itemStartRow, 1, itemRows.length, APP.sheets.ActionItems.length).setNumberFormat('@').setValues(itemRows);
  }
  return { success: true, count: meetingsToImport.length, itemCount: itemRows.length };
}

/**
 * 後端管線自我測試：寫入暫存測試會議、以查詢 API 驗證、最後移除該測試列。
 * 可在 GAS 編輯器選擇此函式並按「執行」，於「執行記錄」查看 [PASS] 或 [FAIL]。
 */
function test_backendPipeline() {
  const date = '2099-12-31';
  const title = '【系統自測】會議寫入管線';
  let meetingId = '';
  try {
    const mockData = JSON.stringify({
      date: date,
      start_time: '09:00',
      end_time: '10:00',
      title: title
    });
    const saved = apiSaveMeeting(mockData);
    meetingId = saved.meeting_id;
    const result = apiGetMeetingsWithItems({ startDate: date, endDate: date });
    const found = result.meetings.some(function (meeting) {
      return meeting.meeting_id === meetingId && meeting.title === title;
    });
    if (!found) throw new Error('查詢結果未找到剛寫入的 meeting_id=' + meetingId);
    Logger.log('[PASS] backend pipeline：資料已寫入並可查詢。meeting_id=' + meetingId);
    return { ok: true, meeting_id: meetingId, message: '[PASS] 寫入與查詢驗證成功。' };
  } catch (error) {
    Logger.log('[FAIL] backend pipeline：' + (error.stack || error.message || error));
    throw new Error('[FAIL] 後端自測失敗：' + (error.message || error));
  } finally {
    // 自測不保留資料，避免污染使用者實際會議清單。
    if (meetingId) {
      try {
        const db = initializeDatabase_();
        const row = findRowByValue_(db.meetings, 1, meetingId);
        if (row) db.meetings.deleteRow(row);
      } catch (cleanupError) {
        Logger.log('[WARN] 自測資料清除失敗：' + (cleanupError.message || cleanupError));
      }
    }
  }
}

/**
 * RSI 閉迴路驗收測試。
 * 測試範圍：資料輸入邊界、0／1／多筆事項關聯、批次匯入、時間格式，以及 Index.html 通訊合約。
 * 測試資料均於 finally 區塊清除；執行記錄應全部顯示 [PASS]。
 */
function run_verification_test() {
  const testDate = '2099-12-30';
  const specialTitle = '【RSI 測試】特殊符號 <>&\"\' / [] #';
  const created = { meetingId: '', itemIds: [], importedMeetingId: '', importedItemIds: [] };
  const trace = [];
  const pass = function (message) {
    const line = '[PASS] ' + message;
    trace.push(line);
    Logger.log(line);
  };

  try {
    const savedMeeting = apiSaveMeeting(JSON.stringify({
      date: testDate,
      start_time: '09:05',
      end_time: '10:10',
      title: specialTitle
    }));
    created.meetingId = savedMeeting.meeting_id;
    assert_(savedMeeting.meeting_id, '有效會議應產生 meeting_id。');

    let found = findMeetingInResult_(apiGetMeetingsWithItems({ startDate: testDate, endDate: testDate }), created.meetingId);
    assert_(found && found.title === specialTitle, '特殊符號標題必須可完整寫入並讀回。');
    assert_(found.items.length === 0, '新會議在尚未新增事項時，items 必須為 0 筆。');
    pass('特殊符號標題與 0 筆改善事項關聯。');

    const firstItem = apiSaveActionItem({
      meeting_id: created.meetingId,
      content: '第一筆改善事項',
      source_person: 'RSI 測試者',
      status: '待處理',
      notes: ''
    });
    created.itemIds.push(firstItem.item_id);
    found = findMeetingInResult_(apiGetMeetingsWithItems({ startDate: testDate, endDate: testDate }), created.meetingId);
    assert_(found.items.length === 1 && found.items[0].item_id === firstItem.item_id, '1 筆改善事項必須正確關聯至會議。');
    pass('1 筆改善事項關聯。');

    const secondItem = apiSaveActionItem(JSON.stringify({
      meeting_id: created.meetingId,
      content: '第二筆改善事項',
      source_person: 'RSI 測試者',
      status: '處理中',
      notes: '多筆關聯測試'
    }));
    created.itemIds.push(secondItem.item_id);
    found = findMeetingInResult_(apiGetMeetingsWithItems({ startDate: testDate, endDate: testDate }), created.meetingId);
    const returnedIds = found.items.map(function (item) { return item.item_id; });
    assert_(found.items.length === 2 && returnedIds.indexOf(firstItem.item_id) !== -1 && returnedIds.indexOf(secondItem.item_id) !== -1,
      '多筆改善事項必須完整且僅關聯至指定會議。');
    pass('多筆改善事項關聯。');

    const globalPendingItems = apiGetAllPendingActionItems();
    const joinedItem = globalPendingItems.filter(function (item) { return item.item_id === firstItem.item_id; })[0];
    assert_(joinedItem && joinedItem.meetingDate === testDate && joinedItem.meetingTitle === specialTitle,
      '全域未結案查詢必須回傳跨月份 Join 後的會議日期與名稱。');
    assert_(globalPendingItems.every(function (item) { return item.status !== '已完成'; }),
      '全域未結案查詢不得包含已完成事項。');
    pass('全域未結案查詢與會議資料 Join。');

    apiUpdateItemStatus(secondItem.item_id, '已完成');
    const globalCompletedItems = apiGetAllCompletedActionItems();
    const joinedCompletedItem = globalCompletedItems.filter(function (item) { return item.item_id === secondItem.item_id; })[0];
    assert_(joinedCompletedItem && joinedCompletedItem.status === '已完成' && joinedCompletedItem.meetingDate === testDate && joinedCompletedItem.meetingTitle === specialTitle,
      '全域已完成查詢必須回傳 Join 後的會議日期、名稱與已完成狀態。');
    pass('全域已完成查詢與會議資料 Join。');

    const importTitle = '【RSI 測試】批次匯入會議';
    const batchResult = apiBatchImportMeetings([{
      date: testDate,
      start_time: '13:00',
      end_time: '14:00',
      title: importTitle,
      items: [{ content: '批次匯入改善事項 A', status: '待處理' }, { content: '批次匯入改善事項 B', status: '處理中' }]
    }]);
    assert_(batchResult.success === true && batchResult.count === 1 && batchResult.itemCount === 2,
      '批次匯入 API 應回傳正確的會議與改善事項數量。');
    const importedMeeting = apiGetMeetingsWithItems({ startDate: testDate, endDate: testDate }).meetings.filter(function (meeting) {
      return meeting.title === importTitle;
    })[0];
    assert_(importedMeeting && importedMeeting.items.length === 2,
      '批次匯入的會議與改善事項必須完整關聯。');
    created.importedMeetingId = importedMeeting.meeting_id;
    created.importedItemIds = importedMeeting.items.map(function (item) { return item.item_id; });
    pass('批次匯入會議與改善事項。');

    assertThrows_(function () {
      apiSaveMeeting({ date: testDate, start_time: '09:00', end_time: '10:00', title: '' });
    }, '空字串標題');
    assertThrows_(function () {
      apiSaveMeeting({ date: testDate, start_time: '9:00', end_time: '10:00', title: '時間格式測試' });
    }, '不符合 HH:mm 的開始時間');
    assertThrows_(function () {
      apiSaveMeeting({ date: testDate, start_time: '09:00', end_time: '24:00', title: '時間格式測試' });
    }, '不符合 HH:mm 的結束時間');
    pass('空標題與 HH:mm 時間格式防呆。');

    verifyFrontendContract_();
    pass('前端純物件傳輸、@click.prevent 與 FailureHandler 合約。');
    Logger.log('[PASS] run_verification_test：全部 ' + trace.length + ' 項驗收條件通過。');
    return { ok: true, passed: trace.length, trace: trace };
  } catch (error) {
    const detail = error.stack || error.message || String(error);
    Logger.log('[FAIL] run_verification_test：' + detail);
    throw new Error('[FAIL] 驗收失敗：' + detail);
  } finally {
    // 先刪除子項目再刪除會議，維持資料關聯完整。
    created.itemIds.concat(created.importedItemIds).forEach(function (itemId) {
      try { apiDeleteActionItem(itemId); } catch (cleanupError) { Logger.log('[WARN] 測試事項清除失敗：' + cleanupError.message); }
    });
    [created.meetingId, created.importedMeetingId].filter(Boolean).forEach(function (meetingId) {
      try {
        const db = initializeDatabase_();
        const row = findRowByValue_(db.meetings, 1, meetingId);
        if (row) db.meetings.deleteRow(row);
      } catch (cleanupError) {
        Logger.log('[WARN] 測試會議清除失敗：' + cleanupError.message);
      }
    });
  }
}

/** 刪除一筆改善事項。 */
function apiDeleteActionItem(itemId) {
  const db = initializeDatabase_();
  const row = findRowByValue_(db.actionItems, 1, itemId);
  if (!row) throw new Error('找不到欲刪除的改善事項。');
  db.actionItems.deleteRow(row);
  return { message: '改善事項已刪除。' };
}

/** 快速更新改善事項狀態。 */
function apiUpdateItemStatus(itemId, newStatus) {
  if (APP.statuses.indexOf(newStatus) === -1) throw new Error('改善事項狀態不正確。');
  const db = initializeDatabase_();
  const row = findRowByValue_(db.actionItems, 1, itemId);
  if (!row) throw new Error('找不到欲更新的改善事項。');
  db.actionItems.getRange(row, 5).setValue(newStatus);
  return { message: '狀態已更新。' };
}

/**
 * 同步會議至執行者主日曆。事件 ID 以 Script Properties 保存，重複同步時更新同一事件。
 */
function apiSyncToGoogleCalendar(meetingId) {
  const db = initializeDatabase_();
  const row = findRowByValue_(db.meetings, 1, meetingId);
  if (!row) throw new Error('找不到欲同步的會議。');
  const values = db.meetings.getRange(row, 1, 1, 7).getDisplayValues()[0];
  const meeting = objectFromRow_(APP.sheets.Meetings, values);
  const start = buildDateTime_(meeting.date, meeting.start_time);
  const end = buildDateTime_(meeting.date, meeting.end_time);
  const key = 'CALENDAR_EVENT_' + meetingId;
  const properties = PropertiesService.getScriptProperties();
  const calendar = CalendarApp.getDefaultCalendar();
  let event = properties.getProperty(key) ? calendar.getEventById(properties.getProperty(key)) : null;
  if (event) {
    event.setTitle(meeting.title).setTime(start, end);
  } else {
    event = calendar.createEvent(meeting.title, start, end, {
      description: '由「會議行事曆與改善事項追蹤系統」同步建立。'
    });
    properties.setProperty(key, event.getId());
  }
  return { message: '已同步至 Google 主日曆。', event_id: event.getId() };
}

/** 建立缺少的分頁及表頭；可手動執行以初始化資料庫。 */
function initSheets() {
  const spreadsheet = getSpreadsheet_();
  const meetings = ensureSheet_(spreadsheet, 'Meetings', APP.sheets.Meetings);
  const actionItems = ensureSheet_(spreadsheet, 'ActionItems', APP.sheets.ActionItems);
  return { meetings: meetings, actionItems: actionItems };
}

function initializeDatabase_() {
  return initSheets();
}

function getSpreadsheet_() {
  // 綁定式專案一定優先使用目前試算表，避免誤將資料寫到其他檔案。
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const configuredId = SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty(APP.spreadsheetIdProperty);
  if (configuredId) return SpreadsheetApp.openById(configuredId);
  throw new Error('找不到資料庫試算表。請在指令碼屬性設定 SPREADSHEET_ID。');
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  const actualHeaders = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  if (actualHeaders.join('|') !== headers.join('|')) {
    if (sheet.getLastRow() > 1) throw new Error('分頁「' + name + '」欄位與系統定義不符，為避免覆寫既有資料已停止初始化。');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e0e7ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readObjects_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues()
    .filter(function (row) { return row[0] !== ''; })
    .map(function (row) { return objectFromRow_(headers, row); });
}

function objectFromRow_(headers, row) {
  return headers.reduce(function (object, header, index) {
    object[header] = row[index] === undefined ? '' : String(row[index]);
    return object;
  }, {});
}

function findRowByValue_(sheet, column, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const values = sheet.getRange(2, column, lastRow - 1, 1).getDisplayValues();
  for (let index = 0; index < values.length; index += 1) {
    if (values[index][0] === String(value)) return index + 2;
  }
  return 0;
}

function isWithinRange_(date, filters) {
  if (filters.month && /^\d{4}-\d{2}$/.test(filters.month) && date.indexOf(filters.month) !== 0) return false;
  if (filters.startDate && date < filters.startDate) return false;
  if (filters.endDate && date > filters.endDate) return false;
  return true;
}

function normalizeDate_(value) {
  const text = String(value || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) throw new Error('日期格式須為 YYYY-MM-DD。');
  const parts = text.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  if (date.getUTCFullYear() !== parts[0] || date.getUTCMonth() !== parts[1] - 1 || date.getUTCDate() !== parts[2]) throw new Error('日期無效。');
  return text;
}

function normalizeTime_(value, label) {
  const text = String(value || '').trim();
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(text)) throw new Error(label + '格式須為 HH:mm。');
  return text;
}

function requiredText_(value, label) {
  const text = String(value || '').trim();
  if (!text) throw new Error('請填寫' + label + '。');
  return text;
}

/** 將 google.script.run 傳入的 JSON 字串或純物件正規化為一般 JavaScript 物件。 */
function parsePayload_(payload, label) {
  let value = payload;
  if (typeof value === 'string') {
    try {
      value = JSON.parse(value);
    } catch (error) {
      throw new Error(label + '不是有效的 JSON 字串：' + error.message);
    }
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(label + '必須是 JSON 字串或物件。');
  }
  return value;
}

/** 將 google.script.run 傳入的 JSON 字串或陣列正規化為一般 JavaScript 陣列。 */
function parseArrayPayload_(payload, label) {
  let value = payload;
  if (typeof value === 'string') {
    try {
      value = JSON.parse(value);
    } catch (error) {
      throw new Error(label + '不是有效的 JSON 字串：' + error.message);
    }
  }
  if (!Array.isArray(value)) throw new Error(label + '必須是 JSON 陣列。');
  return value;
}

function assert_(condition, message) {
  if (!condition) throw new Error(message);
}

function assertThrows_(callback, conditionName) {
  try {
    callback();
  } catch (error) {
    return;
  }
  throw new Error('預期「' + conditionName + '」應被拒絕，但未拋出例外。');
}

function findMeetingInResult_(result, meetingId) {
  return (result.meetings || []).filter(function (meeting) {
    return meeting.meeting_id === meetingId;
  })[0];
}

/** 以 Index.html 原始內容檢核所有 google.script.run 呼叫的必要防呆。 */
function verifyFrontendContract_() {
  const source = HtmlService.createHtmlOutputFromFile('Index').getContent();
  assert_(source.indexOf('JSON.parse(JSON.stringify(this.meetingModal.form))') !== -1,
    '會議表單未在送出前轉換為純 JavaScript 物件。');
  assert_(source.indexOf('@click.prevent="saveMeeting"') !== -1,
    '會議儲存按鈕缺少 @click.prevent。');
  assert_(source.indexOf("viewMode: (() =>") !== -1 && source.indexOf("localStorage.getItem('openItemsViewMode')") !== -1 && source.indexOf("localStorage.setItem('openItemsViewMode', mode)") !== -1,
    '未結案追蹤缺少 Grid／List 檢視偏好保存機制。');
  assert_(source.indexOf('grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4') !== -1 && source.indexOf('flex flex-col gap-2.5') !== -1,
    '未結案追蹤缺少指定的 Grid 或 List 版面結構。');
  assert_(source.indexOf("meetingViewMode: 'list'") !== -1 && source.indexOf('grid grid-cols-7') !== -1 && source.indexOf('Array.from({ length: 42 }') !== -1,
    '會議總覽缺少月曆檢視模式或 42 格月曆矩陣。');
  assert_(source.indexOf('openMeetingModal(null, day.date)') !== -1 && source.indexOf('openMeetingDetails(meeting)') !== -1,
    '月曆方格預填新增或會議膠囊詳細檢視互動未完成。');
  assert_(source.indexOf('globalPendingItems: []') !== -1 && source.indexOf("apiGetAllPendingActionItems") !== -1 && source.indexOf('this.fetchCurrentMonthMeetings(); this.fetchAllPendingItems(); this.fetchAllCompletedItems();') !== -1,
    '前端未建立與月份會議資料解耦的全域未結案載入流程。');
  assert_(source.indexOf("this.globalPendingItems = this.globalPendingItems.filter(pending => pending.item_id !== item.item_id)") !== -1,
    '完成未結案事項後未從全域清單即時移除。');
  assert_(source.indexOf('globalCompletedItems: []') !== -1 && source.indexOf('apiGetAllCompletedActionItems') !== -1,
    '前端未建立全域已完成事項載入流程。');
  assert_(source.indexOf("pendingTimeFilter: 'all'") !== -1 && source.indexOf("completedTimeFilter: '30d'") !== -1 && source.indexOf("customDateRange: { start: '', end: '' }") !== -1,
    '未結案與已完成事項的時間篩選狀態或預設值未獨立設定。');
  assert_(source.indexOf('isWithinTimeFilter') !== -1 && source.indexOf("end.setHours(23, 59, 59, 999)") !== -1,
    '前端缺少自訂日期區間或結束日期全日涵蓋邏輯。');
  assert_(source.indexOf('xlsx@0.18.5/dist/xlsx.full.min.js') !== -1 && source.indexOf('apiBatchImportMeetings') !== -1 && source.indexOf('buildImportPreview') !== -1,
    '前端缺少 Excel／剪貼簿批次匯入與預覽流程。');
  const runCalls = (source.match(/google\.script\.run\s*(?:\.|\n)/g) || []).length;
  const failureHandlers = (source.match(/\.withFailureHandler\s*\(/g) || []).length;
  assert_(runCalls > 0, 'Index.html 未找到 google.script.run 呼叫。');
  assert_(failureHandlers >= runCalls,
    '存在未配置 withFailureHandler 的 google.script.run 非同步呼叫。呼叫數=' + runCalls + '，FailureHandler數=' + failureHandlers);
}

function getChineseWeekday_(dateText) {
  const weekdays = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'];
  const parts = dateText.split('-').map(Number);
  return weekdays[new Date(Date.UTC(parts[0], parts[1] - 1, parts[2])).getUTCDay()];
}

function buildDateTime_(dateText, timeText) {
  const parts = dateText.split('-').map(Number);
  const time = timeText.split(':').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], time[0], time[1], 0);
}
