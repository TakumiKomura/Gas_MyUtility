/**
* CalendeaAppを参照して指定日が営業日かどうか判定する
* @param {string} calendarID カレンダーID
* @param {Date} targetDate 指定日
* @return {boolean} true : 営業日, false : 休業日
*/
this.isWorkDay = (targetDate, calendarID) => {
  let calendar = CalendarApp.getCalendarById(calendarID);
  let options = {
    search: "休館日"
  }

  let holidayEvents = calendar.getEventsForDay(targetDate, options);    
  // 「休館日」というイベントがあったら休業日
  if(holidayEvents.length != 0) return false;

  return true;
};

/**
 * SpreadsheetをExcelファイルに変換してドライブの同じフォルダに保存、Fileを返す
 * @param {SpreadsheetApp.Sheet} spreadsheet_id spreadsheetのid
 * @return {File} new_file 変換後のExcelファイル
 */
this.ss2xlsx = (spreadsheet_id) => {
  let url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id + "/export?format=xlsx";
  let options = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };
  let res = UrlFetchApp.fetch(url, options);
  let new_file;
  if (res.getResponseCode() == 200) {
    console.log("Response Success");
    let ss = SpreadsheetApp.openById(spreadsheet_id);
    new_file = DriveApp.createFile(res.getBlob()).setName(ss.getName() + ".xlsx");
    let parentFolders = DriveApp.getFileById(spreadsheet_id).getParents();
    let folder = parentFolders.next();
    new_file.moveTo(folder); // ルートディレクトリにファイルを作成後、元のファイルと同じフォルダに移動
    console.log("create " + '"' +new_file.getName() + '"' + " file in " + '"' + folder.getName() + '"' + " folder");
  }
  return new_file;
};

/**
 * 指定日時に関数のトリガーをセットする
 * @param {string} funcName トリガーをセットする関数名
 * @param {Date} date 日時(year, month, date [, hour, minute, second, msecond])
 */
this.setTrigger = (funcName, date) => {
  // 同一名の既存のトリガーを削除
  let triggers = ScriptApp.getProjectTriggers();
  for(let trigger of triggers){
    if(trigger.getHandlerFunction() == funcName){
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // トリガーを作成
  ScriptApp.newTrigger(funcName).timeBased().at(date).create();
};
