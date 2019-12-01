 'use strict';
function sendNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("オフィス利用申請");
  const cell = ss.getActiveCell().getA1Notation();
  const columnNameInAlphabet = cell.replace(/\d+/,'');
//通知確認のメールを連想配列で格納
// '依頼者名'：'確認者メール通知先アドレス'←ここを記載すること（フォームから入力された「依頼者」で検索し、下記から依頼者に紐付いたメールアドレスにメールが送られる
  const recipients = { 
    'miii': 'xxxx@gmail.com',
    'aiueo': 'xxxxxx@kinoko.ne.jp'
  };
//更新行列の値を取得（名前が入っているはず）
  const person = sheet.getRange('B'+ sheet.getActiveCell().getRowIndex()).getValue();
//メールの件名と本文を指定
  const subject = '【更新】'+ss.getName();
  const body = person + 'さんがオフィス利用申請フォームより申請しました。\n' + ss.getUrl() + '\nチェックをお願いします。';
//更新された行が通知対象の行を含む場合はメールを送る
 if (person in recipients == true){
    MailApp.sendEmail(recipients[person], subject, body);
  }  
};

function insertLastUpdated(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //現在触っているファイルを取得
  var sheet = ss.getSheetByName('オフィス利用申請'); //対象のシート名を選択
  var currentRow = sheet.getActiveCell().getRow(); //アクティブなセルの行番号を取得
  var currentCol = sheet.getActiveCell().getColumn(); //アクティブなセルの列番号を取得
  var currentCell = sheet.getActiveCell().getValue(); //アクティブなセルの入力値を取得
  var updateRange = sheet.getRange('K' + currentRow); //更新日時を挿入
  Logger.log(updateRange); //更新日の記入
  if(currentRow > 2　&& currentCol == 9) { 
    if(currentCell) {
    updateRange.setValue(new Date());
    }
  }
};
function sendEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("オフィス利用申請");
  var ss_name = SpreadsheetApp.getActive().getSheetByName('オフィス利用申請');

  var name = sheet.getRange(2, 2, ss_name.getLastRow()-1).getValues();
  Logger.log(name); 
  var mailaddress = sheet.getRange(2, 7, ss_name.getLastRow()-1).getValues();
  var cc = sheet.getRange(2, 3, ss_name.getLastRow()-1).getValues();
  var activeRow = sheet.getActiveCell().getRow(); //アクティブなセルの行番号を取得
  var activeCol = sheet.getActiveCell().getColumn(); //アクティブなセルの列番号を取得
  var activeCell = sheet.getActiveCell().getValue(); //アクティブなセルの入力値を取得
  var checkValue = sheet.getDataRange().getValues();
  
  Logger.log(checkValue[activeRow][activeCol]);
  
  //更新された行が通知対象の行を含む場合はメールを送る
 if (activeCol === 8 && checkValue[activeRow][activeCol] === true){
     
  for(var i = 0, l = mailaddress.length; i < l; i++) {
   Logger.log(name[i][0]); 
    MailApp.sendEmail(
    mailaddress[i][0],
      '件名テスト',
      name[i][0] + 'さん' + '\n' + '\n' +
      'おつかれさまです。' + '\n' +
      'オフィス利用申請の確認が実施されました。\n' + '\n' +
      '\n'
      );
    }
  }
 }
