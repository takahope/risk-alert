使用 Apps Script 建立「動態郵件」
如果您希望郵件看起來更專業（像是一個互動式卡片），您可以使用 Google Apps Script 搭配 HTML 電子郵件。

運作邏輯
後端端點：利用 Apps Script 撰寫一個 doGet(e) 或 doPost(e) 函式，並部署為「網頁應用程式 (Web App)」。

HTML 郵件：在郵件中加入帶有參數的 URL 按鈕。

例如：https://script.google.com/.../exec?choice=選項A&user=張三

自動寫入：當使用者點擊按鈕，瀏覽器會觸發 Script，將 choice 和 user 自動寫入 Google Sheet，並回傳一個「感謝完成」的簡單網頁。

核心程式碼範例 (Apps Script)
JavaScript
function doGet(e) {
  var sheet = SpreadsheetApp.openById("您的試算表ID").getActiveSheet();
  var choice = e.parameter.choice;
  var user = e.parameter.user;
  
  // 將資料寫入最後一列
  sheet.appendRow([new Date(), user, choice]);
  
  // 回傳給使用者的確認訊息
  return ContentService.createTextOutput("回覆已成功記錄！您可以關閉此視窗。");
}