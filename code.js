/**
 * @fileoverview Gmail 自動回覆與轉寄處理器 (含儀表板後端)
 * @description 讀取郵件 -> 比對資產 -> 建立草稿 -> 寫入Log -> 提供儀表板資料。
 * @author Google Apps Script Expert
 */

// ==========================================
// 2. 網頁應用程式與設定 API (Web App & Settings)
// ==========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('資安預警即時儀表板')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function getSystemSettings() {
  ensureSettingsSheetExists();
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SETTINGS_SHEET_NAME);
  const values = sheet.getRange("A2:H2").getDisplayValues()[0];
  return {
    scanRead: values[0] === "是",
    autoDraft: values[1] === "是",
    chatNotify: values[2] === "是",
    notInUseSendEmail: values[6] === "是",      // G2: 未使用寄信通知
    processedSendEmail: values[7] === "是"      // H2: 已處理寄信通知
  };
}

function updateSystemSetting(key, isEnabled) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SETTINGS_SHEET_NAME);
  const val = isEnabled ? "是" : "否";
  
  if (key === 'scanRead') sheet.getRange("A2").setValue(val);
  if (key === 'autoDraft') sheet.getRange("B2").setValue(val);
  if (key === 'chatNotify') sheet.getRange("C2").setValue(val);
  if (key === 'notInUseSendEmail') sheet.getRange("G2").setValue(val);   // G2: 未使用寄信通知
  if (key === 'processedSendEmail') sheet.getRange("H2").setValue(val);  // H2: 已處理寄信通知
  
  return { success: true };
}

/**
 * 更新使用狀態（未使用/已處理）
 * @param {number} rowIndex - 資料列索引（從 0 開始）
 * @param {string} usageStatus - 使用狀態（'未使用' 或 '已處理'）
 */
function updateUsageStatus(rowIndex, usageStatus) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!sheet) throw new Error('找不到 Log 工作表');
    
    // rowIndex 是從 0 開始的資料索引，加上標題列後變成實際列號
    const actualRow = rowIndex + 2;
    
    // 讀取該列的資料（Warning Name, Message ID）
    const rowData = sheet.getRange(actualRow, 1, 1, 10).getDisplayValues()[0];
    const warningName = rowData[2];  // C 欄
    const matchedAsset = rowData[3]; // D 欄
    const messageId = rowData[6];    // G 欄
    
    // 更新 I 欄 (使用狀態)
    sheet.getRange(actualRow, 9).setValue(usageStatus);
    
    // 取得操作者資訊
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '未知使用者';
    const userInfo = getUserDisplayName(userEmail);
    
    // 更新 J 欄 (操作者)
    sheet.getRange(actualRow, 10).setValue(userEmail);
    
    // 無論是「未使用」或「已處理」，都將 B 欄狀態從 ALERT 改為 SAFE
    const currentStatus = sheet.getRange(actualRow, 2).getValue();
    if (currentStatus === 'ALERT') {
      sheet.getRange(actualRow, 2).setValue('SAFE');
    }
    
    // 檢查設定
    const settings = getSystemSettings();
    let draftCreated = false;
    let emailSent = false;
    
    if (messageId) {
      // 嘗試透過 Message ID 找到原始郵件
      const originalMessage = findMessageById(messageId);
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      
      if (usageStatus === '未使用') {
        // 未使用情境：優先判斷 notInUseSendEmail
        if (settings.notInUseSendEmail) {
          // 直接寄信回覆寄件者
          emailSent = sendEmailForNotInUse(warningName, matchedAsset, userInfo, originalMessage, settings);
        } else if (settings.autoDraft) {
          // 建立草稿
          draftCreated = createDraftForNotInUse(warningName, matchedAsset, userInfo, originalMessage, settings);
        }
      } else if (usageStatus === '已處理') {
        // 已處理情境：優先判斷 processedSendEmail
        if (settings.processedSendEmail) {
          // 直接寄信回覆寄件者
          emailSent = sendEmailForProcessed(warningName, matchedAsset, userInfo, timestamp, originalMessage, settings);
        } else if (settings.autoDraft) {
          // 建立草稿
          draftCreated = createDraftForProcessed(warningName, matchedAsset, userInfo, timestamp, originalMessage, settings);
        }
      }
      
      // 更新 E 欄 (Action) 加上草稿/寄信資訊
      if (draftCreated || emailSent) {
        const currentAction = sheet.getRange(actualRow, 5).getValue();
        const actionType = emailSent ? '郵件已發送' : '草稿已建立';
        const newAction = currentAction + ` | ${usageStatus}${actionType}`;
        sheet.getRange(actualRow, 5).setValue(newAction);
      }
    }
    
    return { success: true, operator: userEmail, displayName: userInfo.displayName, draftCreated: draftCreated, emailSent: emailSent };
  } catch (e) {
    console.error('更新使用狀態失敗: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 透過 Message ID 查找原始郵件
 * @param {string} messageId - Gmail Message ID (內部 ID)
 * @returns {GmailMessage|null}
 */
function findMessageById(messageId) {
  try {
    if (!messageId) return null;
    
    // [修正] 直接使用 Gmail 內部 ID 取得郵件
    // message.getId() 回傳的是 Gmail 內部 ID，應使用 getMessageById 而非 rfc822msgid 搜尋
    const message = GmailApp.getMessageById(messageId);
    return message || null;
  } catch (e) {
    console.error('查找郵件失敗: ' + e.message);
    return null;
  }
}

function getDashboardData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    // [修改] 讀取所有資料供前端篩選與統計
    const numRows = lastRow - 1;
    
    // 讀取前 10 欄用於儀表板顯示 (Timestamp, Status, WarningName, MatchedAsset, Action, Email Date, Message ID, Has Reply, Not In Use, Operator)
    const values = sheet.getRange(2, 1, numRows, 10).getDisplayValues();
    
    // 依照 Email Date (index 5) 降序排列 (由近到遠)
    values.sort((a, b) => {
      const dateA = a[5] ? new Date(a[5]).getTime() : 0;
      const dateB = b[5] ? new Date(b[5]).getTime() : 0;
      return dateB - dateA; // 降序
    });
    
    return values;
  } catch (e) {
    console.error("獲取儀表板資料失敗: " + e.message);
    return [];
  }
}

/**
 * 取得指定紀錄的郵件內容（漸進式載入）
 * 如果 K 欄已有內容則直接回傳，否則透過 Message ID 從 Gmail 讀取並存入 K 欄
 * @param {number} rowIndex - 資料列索引（從 0 開始，對應前端 filteredData 的索引）
 * @param {string} msgId - Message ID（從前端傳入）
 * @returns {Object} { success: boolean, content: string, error?: string }
 */
function getEmailContent(rowIndex, msgId) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!sheet) throw new Error('找不到 Log 工作表');
    
    // 由於前端資料經過排序，我們需要透過 msgId 找到正確的列
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('沒有資料');
    
    // 讀取所有 Message ID (G 欄) 和 Email Content (K 欄)
    const allData = sheet.getRange(2, 7, lastRow - 1, 5).getValues(); // G 到 K 欄
    
    // 找到對應 msgId 的列
    let targetRow = -1;
    for (let i = 0; i < allData.length; i++) {
      if (allData[i][0] === msgId) { // G 欄是 Message ID
        targetRow = i + 2; // 實際列號 (加上標題列)
        
        // 檢查 K 欄 (index 4) 是否已有內容
        const existingContent = allData[i][4];
        if (existingContent && existingContent.trim() !== '') {
          return { success: true, content: existingContent };
        }
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error('找不到對應的紀錄');
    }
    
    // K 欄為空，需要從 Gmail 讀取
    if (!msgId) {
      return { success: false, content: '', error: '此紀錄沒有關聯的郵件 ID' };
    }
    
    // 透過 Message ID 查找郵件
    const message = findMessageById(msgId);
    if (!message) {
      return { success: false, content: '', error: '無法找到對應的郵件，可能已被刪除' };
    }
    
    // 取得郵件內容
    const emailContent = message.getPlainBody();
    const contentToSave = emailContent.substring(0, 10000); // 限制最大 10000 字
    
    // 儲存到 K 欄
    sheet.getRange(targetRow, 11).setValue(contentToSave);
    
    return { success: true, content: contentToSave };
    
  } catch (e) {
    console.error('取得郵件內容失敗: ' + e.message);
    return { success: false, content: '', error: e.message };
  }
}

// ==========================================
// 3. 主控制器 (Main Controller)
// ==========================================

function processIncomingEmails() {
  // 1. 讀取當前設定
  const settings = getSystemSettings();
  
  // 2. 構建搜尋語法 (支援多個寄件者和多個標題關鍵字)
  const fromClause = CONFIG.SENDER_B_EMAILS.length === 1
    ? `from:${CONFIG.SENDER_B_EMAILS[0]}`
    : `{${CONFIG.SENDER_B_EMAILS.map(e => `from:${e}`).join(' ')}}`;
  
  const subjectClause = CONFIG.SUBJECT_KEYWORDS.length === 1
    ? `subject:"${CONFIG.SUBJECT_KEYWORDS[0]}"`
    : `{${CONFIG.SUBJECT_KEYWORDS.map(k => `subject:"${k}"`).join(' ')}}`;
  
  let query = `${fromClause} ${subjectClause}`;
  
  // 如果設定 A2="否" (預設)，則只抓未讀；若 A2="是"，則不加限制(會掃描所有信，靠 MessageID 去重)
  if (!settings.scanRead) {
    query += ` is:unread`;
  }

  const threads = GmailApp.search(query);
  if (threads.length === 0) {
    console.log("目前沒有符合條件的信件。");
    // 回傳結果給前端顯示 (可選)
    return "無符合條件信件";
  }

  ensureLogSheetExists();
  const assetList = fetchComparisonData();
  const processedMessageIds = fetchProcessedMessageIds();
  let processCount = 0;

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      // 去重檢查
      if (processedMessageIds.has(message.getId())) return;
      
      // 雙重防護：如果設定只掃未讀，但信已讀，則跳過
      if (!settings.scanRead && !message.isUnread()) return;

      processSingleMessage(message, assetList, settings);
      processCount++;
      
      if (message.isUnread()) {
        GmailApp.markMessageRead(message);
      }
    });
  });
  
  return `掃描完成，處理了 ${processCount} 封新信件`;
}

function processSingleMessage(message, assetList, settings) {
  const body = message.getPlainBody();
  const warningInfo = extractWarningInfo(body);
  const warningName = warningInfo.displayName;
  const matchFields = warningInfo.matchFields;
  const msgId = message.getId();
  
  // [新增] 提取信件收到時間並格式化
  const emailDate = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  
  // [新增] 檢查討論串資訊
  const thread = message.getThread();
  const allMessages = thread.getMessages();
  const threadMessageCount = allMessages.length;
  const hasReply = threadMessageCount > 1 ? '有回覆' : '';
  
  // [新增] 檢查最新討論串內容是否包含「未使用」
  let notInUse = '';
  if (threadMessageCount > 1) {
    // 取得最新的訊息（最後一封）
    const latestMessage = allMessages[allMessages.length - 1];
    const latestSender = latestMessage.getFrom();
    
    // 檢查回覆者是否為原始篩選的寄件者（若是則略過）
    const isFromOriginalSender = CONFIG.SENDER_B_EMAILS.some(email => 
      latestSender.toLowerCase().includes(email.toLowerCase())
    );
    
    // 只有當回覆者不是原始寄件者時，才檢查「未使用」
    if (!isFromOriginalSender) {
      const latestBody = latestMessage.getPlainBody();
      if (latestBody.includes('未使用')) {
        notInUse = '未使用';
      }
    }
  }
  
  // [新增] 特殊處理：如果包含「DN 與 IP 黑名單」關鍵字，直接使用固定值
  if (warningInfo.isBlacklistAlert && !warningName) {
    let actionLog = '僅紀錄 (自動草稿已關閉)';
    if (settings.autoDraft) {
      // 建立通知草稿給 Person A
      const fixedWarningName = 'DN 與 IP 黑名單更新';
      const fixedAsset = 'DN 與 IP 黑名單';
      createDraftForPersonA(fixedWarningName, fixedAsset, message, settings);
      actionLog = '已建立通知草稿';
    } else if (settings.chatNotify) {
      sendToChat(`🚨 **[DN 與 IP 黑名單更新]**\n已偵測到黑名單更新通知`);
    }
    logExecutionResult('ALERT', 'DN 與 IP 黑名單更新', 'DN 與 IP 黑名單', actionLog, msgId, emailDate, hasReply, notInUse, '');
    return;
  }
  
  if (!warningName) {
    const errorMsg = `無法提取警訊名稱或漏洞說明`;
    // [修改] 不自動存入郵件內容，改為用戶點擊時才載入
    logExecutionResult('ERROR', '解析失敗', 'N/A', errorMsg, msgId, emailDate, hasReply, notInUse, '');
    if (settings.chatNotify) sendToChat(`⚠️ 錯誤報告：${errorMsg}`);
    return; 
  }

  // 搜尋所有命中的資產（比對所有提取的欄位：警訊名稱、漏洞說明、內容說明、影響平台、影響等級）
  let matchedAssets = assetList.filter(asset => 
    matchFields.some(field => field.toLowerCase().includes(asset.toLowerCase()))
  );
  
  // [新增] 去除重複的資產（不區分大小寫）
  const uniqueAssets = [...new Set(matchedAssets.map(a => a.toLowerCase()))]
    .map(lowerAsset => matchedAssets.find(a => a.toLowerCase() === lowerAsset));

  if (uniqueAssets.length > 0) {
    // [修改] 將所有命中資產合併為字串
    const matchedAssetStr = uniqueAssets.join(', ');
    
    let actionLog = '僅紀錄';
    
    // 優先判斷 notInUseSendEmail（直接寄信）
    if (settings.notInUseSendEmail) {
      sendEmailForPersonA(warningName, matchedAssetStr, message, settings);
      actionLog = '已直接寄信給 Person A';
    } else if (settings.autoDraft) {
      // 其次判斷 autoDraft（建立草稿）
      createDraftForPersonA(warningName, matchedAssetStr, message, settings);
      actionLog = '已建立通知草稿';
    } else if (settings.chatNotify) {
       sendToChat(`🚨 **[資產命中] (草稿功能未啟用)**\n偵測資產：${matchedAssetStr}\n警訊資訊：${warningName}`);
    }
    // [修改] 不自動存入郵件內容，改為用戶點擊時才載入
    logExecutionResult('ALERT', warningName, matchedAssetStr, actionLog, msgId, emailDate, hasReply, notInUse, '');
  } else {
    // 無命中資產：直接標記為 SAFE 和 未使用
    let actionLog = '僅紀錄';
    
    // 優先判斷 notInUseSendEmail（直接寄信回覆）
    if (settings.notInUseSendEmail) {
      sendEmailReplyToSenderB(warningName, message, settings);
      actionLog = '已直接寄信回覆';
    } else if (settings.autoDraft) {
      // 其次判斷 autoDraft（建立回覆草稿）
      createDraftReplyToSenderB(warningName, message, settings);
      actionLog = '已建立回覆草稿';
    }
    // [修改] 不自動存入郵件內容，改為用戶點擊時才載入
    logExecutionResult('SAFE', warningName, '無相關資產', actionLog, msgId, emailDate, hasReply, '未使用', '');
  }
}

// ==========================================
// 4. 資料提取與資料庫服務
// ==========================================

function extractWarningInfo(text) {
  // 警訊名稱
  const nameRegex = /警訊名稱[：:]\s*(.+)/i;
  // 漏洞說明
  const descRegex = /漏洞說明[：:]\s*(.+)/i;
  // [新增] 內容說明
  const contentRegex = /內容說明[：:]\s*(.+)/i;
  // [新增] 影響平台
  const platformRegex = /影響平台[：:]\s*(.+)/i;
  // [新增] 影響等級
  const levelRegex = /影響等級[：:]\s*(.+)/i;
  
  const nameMatch = text.match(nameRegex);
  const descMatch = text.match(descRegex);
  const contentMatch = text.match(contentRegex);
  const platformMatch = text.match(platformRegex);
  const levelMatch = text.match(levelRegex);
  
  const name = (nameMatch && nameMatch[1]) ? nameMatch[1].trim() : null;
  const desc = (descMatch && descMatch[1]) ? descMatch[1].trim() : null;
  const content = (contentMatch && contentMatch[1]) ? contentMatch[1].trim() : null;
  const platform = (platformMatch && platformMatch[1]) ? platformMatch[1].trim() : null;
  const level = (levelMatch && levelMatch[1]) ? levelMatch[1].trim() : null;
  
  // [新增] 檢查是否包含「DN 與 IP 黑名單」關鍵字（不需要冒號格式）
  const hasBlacklistKeyword = /DN\s*與\s*IP\s*黑名單/i.test(text);
  
  // 組合顯示結果
  let displayName = '';
  
  // 優先使用警訊名稱，其次是漏洞說明、內容說明
  if (name) {
    displayName = name;
  } else if (desc) {
    displayName = desc;
  } else if (content) {
    displayName = content;
  }
  
  // 如果有結果，附加額外資訊
  if (displayName) {
    const extras = [];
    if (desc && name) extras.push(`說明: ${desc}`);
    if (content && !displayName.includes(content)) extras.push(`內容: ${content}`);
    if (platform) extras.push(`平台: ${platform}`);
    if (level) extras.push(`等級: ${level}`);
    
    if (extras.length > 0) {
      displayName += ` (${extras.join(' | ')})`;
    }
  }
  
  // 返回物件：包含顯示名稱、主要匹配欄位和是否為黑名單警示
  return {
    displayName: displayName || null,
    matchFields: [name, desc, content, platform, level].filter(Boolean),
    isBlacklistAlert: hasBlacklistKeyword  // [新增] 標記是否包含黑名單關鍵字
  };
}

/**
 * [修改] 從 Google Sheet 獲取資產清單 (擴大為 A, B, C 三欄)
 */
function fetchComparisonData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error("找不到資產清單工作表");
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return [];
    
    // 修改處：讀取 3 欄 (A, B, C)，從第 1 欄開始，讀 3 欄寬度
    return sheet.getRange(1, 1, lastRow, 3)
      .getValues()
      .flat() // 將 [[A1,B1,C1], [A2,B2,C2]] 攤平成 [A1,B1,C1,A2...]
      .map(String)
      .map(s => s.trim())
      .filter(s => s.length > 0); // 過濾空字串
  } catch (e) {
    console.error("讀取試算表失敗: " + e.message);
    return [];
  }
}

function fetchProcessedMessageIds() {
  const ids = new Set();
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!sheet) return ids;
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return ids;
    // [修改] Message ID 現在位於第 7 欄 (G欄)，因為第 6 欄變成了 Email Date
    const data = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
    data.flat().forEach(id => { if(id) ids.add(String(id)); });
  } catch (e) {
    console.error("讀取 Processed IDs 失敗: " + e.message);
  }
  return ids;
}

function ensureSettingsSheetExists() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let settingSheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET_NAME);
  if (!settingSheet) {
    settingSheet = ss.insertSheet(CONFIG.SETTINGS_SHEET_NAME);
    settingSheet.appendRow(['掃描已讀信件 (A2)', '開啟自動草稿 (B2)', '開啟Chat通知 (C2)', '授權使用者 (D欄)', '使用者姓名 (E欄)', '保留欄位 (F欄)', '未使用寄信通知 (G2)', '已處理寄信通知 (H2)']);
    settingSheet.appendRow(['否', '是', '是', '', '', '', '否', '否']);
  }
}

/**
 * 取得使用者顯示名稱（姓名 + email）
 * @param {string} email - 使用者 email
 * @returns {Object} { displayName: string, name: string, email: string }
 */
function getUserDisplayName(email) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { displayName: email, name: '', email: email };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { displayName: email, name: '', email: email };
    }
    
    // 讀取 D 欄 (email) 和 E 欄 (姓名)
    const userData = sheet.getRange(2, 4, lastRow - 1, 2).getDisplayValues();
    
    for (let i = 0; i < userData.length; i++) {
      const userEmail = userData[i][0].trim().toLowerCase();
      const userName = userData[i][1].trim();
      
      if (userEmail === email.toLowerCase() && userName) {
        return {
          displayName: `${userName} (${email})`,
          name: userName,
          email: email
        };
      }
    }
    
    return { displayName: email, name: '', email: email };
  } catch (e) {
    console.error('取得使用者名稱失敗: ' + e.message);
    return { displayName: email, name: '', email: email };
  }
}

/**
 * 檢查當前使用者是否有權限
 * @returns {Object} { authorized: boolean, email: string, message: string }
 */
function checkUserAuthorization() {
  try {
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    
    if (!userEmail) {
      return { authorized: false, email: '', message: '無法取得使用者資訊' };
    }
    
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { authorized: false, email: userEmail, message: '無法讀取設定' };
    }
    
    // 讀取 D 欄所有授權使用者
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { authorized: false, email: userEmail, message: '尚未設定授權使用者' };
    }
    
    const authorizedUsers = sheet.getRange(2, 4, lastRow - 1, 1).getDisplayValues()
      .flat()
      .map(email => email.trim().toLowerCase())
      .filter(email => email.length > 0);
    
    // 如果 D 欄為空，允許所有使用者
    if (authorizedUsers.length === 0) {
      return { authorized: true, email: userEmail, message: '歡迎使用' };
    }
    
    const isAuthorized = authorizedUsers.includes(userEmail.toLowerCase());
    
    return {
      authorized: isAuthorized,
      email: userEmail,
      message: isAuthorized ? '歡迎使用' : '您沒有權限使用此系統'
    };
  } catch (e) {
    console.error('權限檢查失敗: ' + e.message);
    return { authorized: false, email: '', message: '權限檢查失敗: ' + e.message };
  }
}

function ensureLogSheetExists() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  
  // 標準標題列
  const headers = ['Timestamp', 'Status', 'Warning Name', 'Matched Asset', 'Action', 'Email Date', 'Message ID', 'Has Reply', 'Not In Use', 'Operator', 'Email Content'];
  
  if (!sheet) {
    // 工作表不存在，建立新的
    sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    // 工作表已存在，檢查第一列是否有標題
    const firstRowValue = sheet.getRange(1, 1).getValue();
    
    // 如果第一列第一格為空或不是 'Timestamp'，則寫入標題
    if (!firstRowValue || firstRowValue !== 'Timestamp') {
      // 在第一列插入新行作為標題
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
}

function logExecutionResult(status, warningName, asset, action, msgId, emailDate, hasReply = '', notInUse = '', emailContent = '') {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (sheet) {
      const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      // [修改] 寫入順序加入 hasReply, notInUse 和 emailContent (K欄)
      sheet.appendRow([time, status, warningName, asset, action, emailDate, msgId, hasReply, notInUse, '', emailContent]);
    }
  } catch (e) {
    console.error("寫入 Log 失敗: " + e.message);
  }
}

// ==========================================
// 5. 通知與草稿服務
// ==========================================

function createDraftForPersonA(warningName, matchedAsset, originalMessage, settings) {
  const subject = `[資產風險警示] 發現內部資產 ${matchedAsset} 相關漏洞`;
  // [修復] 恢復完整的信件內容
  const body = `
    親愛的 A：
    
    系統檢測到最新的漏洞預警通報內容與資訊資產名稱 "${matchedAsset}" 有相關性。
    警訊名稱：${warningName}
    
    請儘速確認並評估影響範圍。
    
    原始信件內容摘要：
    ${originalMessage.getPlainBody().substring(0, 300)}...
  `;
  
  GmailApp.createDraft(CONFIG.PERSON_A_EMAIL, subject, body);
  
  if (settings.chatNotify) {
    sendToChat(`🚨 **[資產命中] 已建立通知草稿**\n偵測資產：${matchedAsset}\n警訊名稱：${warningName}`);
  }
}

function createDraftReplyToSenderB(warningName, originalMessage, settings) {
  // [修復] 恢復完整的信件內容
  const replyBody = `
    您好，
    
    已收到漏洞預警通知：
    "${warningName}"
    
    經確認，無相關軟硬體資產，無需處理。
    感謝通知。
  `;
  
  originalMessage.getThread().createDraftReply(replyBody);
  
  if (settings.chatNotify) {
    sendToChat(`✅ **[無相關資產] 已建立回覆草稿**\n警訊名稱：${warningName}`);
  }
}

// ==========================================
// 5.1 直接寄信服務（未使用寄信通知功能）
// ==========================================

/**
 * 直接寄信給 Person A（資產命中時）
 * @param {string} warningName - 警訊名稱
 * @param {string} matchedAsset - 命中資產
 * @param {GmailMessage} originalMessage - 原始郵件
 * @param {Object} settings - 系統設定
 */
function sendEmailForPersonA(warningName, matchedAsset, originalMessage, settings) {
  try {
    const subject = `[資產風險警示] 發現內部資產 ${matchedAsset} 相關漏洞`;
    const body = `
親愛的 A：

系統檢測到最新的漏洞預警通報內容與資訊資產名稱 "${matchedAsset}" 有相關性。
警訊名稱：${warningName}

請儘速確認並評估影響範圍。

原始信件內容摘要：
${originalMessage.getPlainBody().substring(0, 300)}...

此郵件由系統自動發送。
    `.trim();
    
    GmailApp.sendEmail(CONFIG.PERSON_A_EMAIL, subject, body);
    
    if (settings.chatNotify) {
      sendToChat(`📧 **[資產命中] 已直接寄信給 Person A**\n偵測資產：${matchedAsset}\n警訊名稱：${warningName}`);
    }
    
    return true;
  } catch (e) {
    console.error('寄信給 Person A 失敗: ' + e.message);
    return false;
  }
}

/**
 * 直接回覆寄件者（無相關資產時）
 * @param {string} warningName - 警訊名稱
 * @param {GmailMessage} originalMessage - 原始郵件
 * @param {Object} settings - 系統設定
 */
function sendEmailReplyToSenderB(warningName, originalMessage, settings) {
  try {
    const replyBody = `
您好，

已收到漏洞預警通知：
「${warningName}」

經確認，無相關軟硬體資產，無需處理。
感謝通知。

此郵件由系統自動發送。
    `.trim();
    
    originalMessage.reply(replyBody);
    
    if (settings.chatNotify) {
      sendToChat(`📧 **[無相關資產] 已直接寄信回覆**\n警訊名稱：${warningName}`);
    }
    
    return true;
  } catch (e) {
    console.error('回覆寄件者失敗: ' + e.message);
    return false;
  }
}

/**
 * 直接寄信回覆（未使用情境 - 手動操作）
 * @param {string} warningName - 警訊名稱
 * @param {string} matchedAsset - 命中資產
 * @param {Object} userInfo - 操作者資訊
 * @param {GmailMessage|null} originalMessage - 原始郵件
 * @param {Object} settings - 系統設定
 * @returns {boolean} 是否成功寄信
 */
function sendEmailForNotInUse(warningName, matchedAsset, userInfo, originalMessage, settings) {
  try {
    const replyBody = `
您好，

已收到漏洞預警通知：
「${warningName}」

經人工確認，相關資產「${matchedAsset}」無需處理。
感謝通知。

處理人員：${userInfo.displayName}

此郵件由系統自動發送。
    `.trim();
    
    if (originalMessage) {
      // 直接回覆原始郵件
      originalMessage.reply(replyBody);
    } else {
      // 無法找到原始郵件，發送獨立郵件
      const subject = `[無需處理] ${warningName}`;
      GmailApp.sendEmail(CONFIG.SENDER_B_EMAILS[0] || '', subject, replyBody);
    }
    
    if (settings.chatNotify) {
      sendToChat(`📧 **[未使用] 已直接寄信回覆**\n警訊：${warningName}\n資產：${matchedAsset}\n處理者：${userInfo.displayName}`);
    }
    
    return true;
  } catch (e) {
    console.error('寄信回覆失敗: ' + e.message);
    return false;
  }
}

/**
 * 直接寄信回覆（已處理情境 - 手動操作）
 * @param {string} warningName - 警訊名稱
 * @param {string} matchedAsset - 命中資產
 * @param {Object} userInfo - 操作者資訊
 * @param {string} timestamp - 處理時間
 * @param {GmailMessage|null} originalMessage - 原始郵件
 * @param {Object} settings - 系統設定
 * @returns {boolean} 是否成功寄信
 */
function sendEmailForProcessed(warningName, matchedAsset, userInfo, timestamp, originalMessage, settings) {
  try {
    const replyBody = `
您好，

關於漏洞預警通知：
「${warningName}」

經評估確認影響範圍，相關資產「${matchedAsset}」已完成必要之風險處置措施。

處理人員：${userInfo.displayName}
處理時間：${timestamp}

如有任何問題，請隨時聯繫。

此郵件由系統自動發送。
    `.trim();
    
    if (originalMessage) {
      // 直接回覆原始郵件
      originalMessage.reply(replyBody);
    } else {
      // 無法找到原始郵件，發送獨立郵件
      const subject = `[已處理] ${warningName}`;
      GmailApp.sendEmail(CONFIG.SENDER_B_EMAILS[0] || '', subject, replyBody);
    }
    
    if (settings.chatNotify) {
      sendToChat(`📧 **[已處理] 已直接寄信回覆**\n警訊：${warningName}\n資產：${matchedAsset}\n處理者：${userInfo.displayName}\n時間：${timestamp}`);
    }
    
    return true;
  } catch (e) {
    console.error('已處理寄信回覆失敗: ' + e.message);
    return false;
  }
}

/**
 * 建立「未使用」情境的回覆草稿
 * @param {string} warningName - 警訊名稱
 * @param {string} matchedAsset - 命中資產
 * @param {Object} userInfo - 操作者資訊
 * @param {GmailMessage|null} originalMessage - 原始郵件
 * @param {Object} settings - 系統設定
 * @returns {boolean} 是否成功建立草稿
 */
function createDraftForNotInUse(warningName, matchedAsset, userInfo, originalMessage, settings) {
  try {
    const replyBody = `
您好，

已收到漏洞預警通知：
「${warningName}」

經人工確認，相關資產「${matchedAsset}」無需處理。
感謝通知。

處理人員：${userInfo.displayName}
    `.trim();
    
    if (originalMessage) {
      // 在原始執行緒建立回覆草稿
      originalMessage.getThread().createDraftReply(replyBody);
    } else {
      // 無法找到原始郵件，建立獨立草稿
      const subject = `[無需處理] ${warningName}`;
      GmailApp.createDraft(CONFIG.SENDER_B_EMAILS[0] || '', subject, replyBody);
    }
    
    if (settings.chatNotify) {
      sendToChat(`🚫 **[未使用] 已建立回覆草稿**\n警訊：${warningName}\n資產：${matchedAsset}\n處理者：${userInfo.displayName}`);
    }
    
    return true;
  } catch (e) {
    console.error('建立未使用草稿失敗: ' + e.message);
    return false;
  }
}

/**
 * 建立「已處理」情境的通知草稿
 * @param {string} warningName - 警訊名稱
 * @param {string} matchedAsset - 命中資產
 * @param {Object} userInfo - 操作者資訊
 * @param {string} timestamp - 處理時間
 * @param {GmailMessage|null} originalMessage - 原始郵件
 * @param {Object} settings - 系統設定
 * @returns {boolean} 是否成功建立草稿
 */
function createDraftForProcessed(warningName, matchedAsset, userInfo, timestamp, originalMessage, settings) {
  try {
    const replyBody = `
您好，

關於漏洞預警通知：
「${warningName}」

經評估確認影響範圍，相關資產「${matchedAsset}」已完成必要之風險處置措施。

處理人員：${userInfo.displayName}
處理時間：${timestamp}

如有任何問題，請隨時聯繫。
    `.trim();
    
    if (originalMessage) {
      // 在原始執行緒建立回覆草稿
      originalMessage.getThread().createDraftReply(replyBody);
    } else {
      // 無法找到原始郵件，建立獨立草稿
      const subject = `[已處理] ${warningName}`;
      GmailApp.createDraft(CONFIG.SENDER_B_EMAILS[0] || '', subject, replyBody);
    }
    
    if (settings.chatNotify) {
      sendToChat(`✅ **[已處理] 已建立通知草稿**\n警訊：${warningName}\n資產：${matchedAsset}\n處理者：${userInfo.displayName}\n時間：${timestamp}`);
    }
    
    return true;
  } catch (e) {
    console.error('建立已處理草稿失敗: ' + e.message);
    return false;
  }
}
function sendToChat(text) {
  if (!CONFIG.GOOGLE_CHAT_WEBHOOK_URL || CONFIG.GOOGLE_CHAT_WEBHOOK_URL.includes('YOUR_KEY')) {
    console.log("模擬 Chat 通知: " + text);
    return;
  }
  try {
    UrlFetchApp.fetch(CONFIG.GOOGLE_CHAT_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: text }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error("Chat 通知失敗: " + e.message);
  }
}