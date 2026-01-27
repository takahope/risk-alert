// ==========================================
// 1. 全局設定 (Configuration)
// ==========================================
const CONFIG = {
  // [修改] 支援多個寄件者信箱 (陣列格式)
  SENDER_B_EMAILS: [
    'actix@siu.tw',
    // 可在此新增更多信箱地址
  ],
  
  // [新增] 篩選信件標題的關鍵字
  SUBJECT_KEYWORD: '【漏洞預警】',
  
  PERSON_A_EMAIL: 'eric@du.tw',
  SPREADSHEET_ID: '1uRyOZecs0GoQ7QMI', 
  SHEET_NAME: 'Sheet1', // 資產清單
  LOG_SHEET_NAME: 'SystemLogs', // [新增] 用於儲存比對紀錄的工作表名稱
  SETTINGS_SHEET_NAME: 'Settings', // [新增] 設定表
  DATA_COLUMN_INDEX: 0, 
  
  GOOGLE_CHAT_WEBHOOK_URL: 'https://chat.googleapis.coWEfRq3CPzqKqqsHEWtgS9oggV4qis8',
  
  // [修改] 搜尋語法：支援多個寄件者，使用 Gmail 的 {from:x from:y} 語法
  getSearchQuery: function() {
    const fromClause = this.SENDER_B_EMAILS.length === 1
      ? `from:${this.SENDER_B_EMAILS[0]}`
      : `{${this.SENDER_B_EMAILS.map(e => `from:${e}`).join(' ')}}`;
    return `${fromClause} subject:"${this.SUBJECT_KEYWORD}" is:unread`;
  }
};