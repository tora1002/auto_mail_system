// グローバル変数
var ss = SpreadsheetApp.getActiveSpreadsheet();

// 実行メニューを作成
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu("メールメニュー");
    menu.addItem("領収書不備", "sendReceiptMistake");
    menu.addItem("収支確認", "sendBalanceCheck");
    menu.addItem("日経テレコン利用ID・PW変更", "sendNikkei");
    menu.addItem("未着・不備請求書", "send");
    menu.addItem("検収チェックシート捺印", "send");
    menu.addItem("未着・不備請求書", "send");
    menu.addItem("検修書提出", "send");
    menu.addToUi();
}

function sendReceiptMistake(){
    var sheet = ss.getSheetByName("領収書不備");
    var startRow = 8;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();

    var strFrom = sheet.getRange(1,2).getValue();

    var docBaseID = sheet.getRange(2,2).getValue();
    var docVariableID = sheet.getRange(3,2).getValue();

    var accountingMonth = sheet.getRange(4,2).getValue();
    var strFixedSubject = sheet.getRange(5,2).getValue();

    // テンプレートテキストの取得  
    var docBaseTemplate = DocumentApp.openById(docBaseID);
    var docVariableTemplate = DocumentApp.openById(docVariableID);
    var strBaseTemplate = docBaseTemplate.getBody().getText();
    var strVariableTemplate = docVariableTemplate.getBody().getText();

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        row.rowNumber = i + startRow;

        // Result列がブランクであれば処理を実行    
        if (!row[11]) { 
            var result = "";

            try
            {
                var strTo = row[0];
                var strCc = row[1];
                var strDestinationSubject = row[2];

                // メールの件名を作成
                var strSubject = "【" + accountingMonth + "月経費】" + strFixedSubject + "（" + strDestinationSubject + "）";

                var options = {};
                options.cc = strCc;
                options.from = strFrom;

                // メールのbase部分の変数を取得
                var strVal1 = row[3];
                var strVal2 = row[4];
                var strVal3 = row[5];
                var strVal4 = row[6];
                
                // メールのbase部分の変数を置換
                var strBody = strBaseTemplate.replace("\{VALUE1\}",strVal1).replace("\{VALUE2\}",strVal2).replace("\{VALUE3\}",strVal3).replace("\{VALUE4\}",strVal4); 

                // メールのvariable部分の変数を取得
                var strVal5 = row[7];
                var strVal6 = row[8];
                var strVal7 = row[9];
                var strVal8 = row[10];

                var strVariable = strVariableTemplate.replace("\{VALUE5\}",strVal5).replace("\{VALUE6\}",strVal6).replace("\{VALUE7\}",strVal7).replace("\{VALUE8\}",strVal8); 

                while (data[i+1] != undefined && strTo == data[i+1][0]) {
                    var strVal5 = data[i+1][7];
                    var strVal6 = data[i+1][8];
                    var strVal7 = data[i+1][9];
                    var strVal8 = data[i+1][10];

                    var strVariable = strVariable + strVariableTemplate.replace("\{VALUE5\}",strVal5).replace("\{VALUE6\}",strVal6).replace("\{VALUE7\}",strVal7).replace("\{VALUE8\}",strVal8);

                    i = i + 1;
                }
                
                // メールのvariable部分の変数を置換
                var strBody = strBody.replace("\{VALUE_variable\}",strVariable); 

                // メール送信実行       
                GmailApp.sendEmail(strTo,strSubject,strBody,options);

                result = "Success"; 
            }catch(e){
                result = "Error:" + e;
            }

            // 実行結果をResult列にセット
            sheet.getRange(row.rowNumber, lastColum).setValue(result); 
        }
    }  
}

function sendBalanceCheck(){
    var sheet = ss.getSheetByName("収支確認");
    var startRow = 7;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();

    var strFrom = sheet.getRange(1,2).getValue();

    var docBaseID = sheet.getRange(2,2).getValue();
    var docVariableID = sheet.getRange(3,2).getValue();

    var strSubject = sheet.getRange(4,2).getValue();

    // テンプレートテキストの取得  
    var docBaseTemplate = DocumentApp.openById(docBaseID);
    var docVariableTemplate = DocumentApp.openById(docVariableID);
    var strBaseTemplate = docBaseTemplate.getBody().getText();
    var strVariableTemplate = docVariableTemplate.getBody().getText();

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        row.rowNumber = i + startRow;

        // Result列がブランクであれば処理を実行    
        if (!row[9]) { 
            var result = "";

            try
            {
                var strTo = row[0];
                var strCc = row[1];

                var options = {};
                options.cc = strCc;
                options.from = strFrom;

                // メールのbase部分の変数を取得
                var strVal1 = row[2];
                var strVal2 = row[3];
                var strVal3 = row[4];
                
                // メールのbase部分の変数を置換
                var strBody = strBaseTemplate.replace("\{VALUE1\}",strVal1).replace("\{VALUE2\}",strVal2).replace("\{VALUE3\}",strVal3); 

                // メールのvariable部分の変数を取得
                var strVal4 = row[5];
                var strVal5 = row[6];
                var strVal6 = row[7];
                var strVal7 = row[8];

                var strVariable = strVariableTemplate.replace("\{VALUE4\}",strVal4).replace("\{VALUE5\}",strVal5).replace("\{VALUE6\}",strVal6).replace("\{VALUE7\}",strVal7); 

                while (data[i+1] != undefined && strTo == data[i+1][0]) {
                    var strVal4 = data[i+1][5];
                    var strVal5 = data[i+1][6];
                    var strVal6 = data[i+1][7];
                    var strVal7 = data[i+1][8];

                    var strVariable = strVariable + strVariableTemplate.replace("\{VALUE4\}",strVal4).replace("\{VALUE5\}",strVal5).replace("\{VALUE6\}",strVal6).replace("\{VALUE7\}",strVal7);

                    i = i + 1;
                }
                
                // メールのvariable部分の変数を置換
                var strBody = strBody.replace("\{VALUE_variable\}",strVariable); 

                // メール送信実行       
                GmailApp.sendEmail(strTo,strSubject,strBody,options);

                result = "Success"; 
            }catch(e){
                result = "Error:" + e;
            }

            // 実行結果をResult列にセット
            sheet.getRange(row.rowNumber, lastColum).setValue(result); 
        }
    }  
}

function sendNikkei(){
    var sheet = ss.getSheetByName("日経テレコン利用ID・PW変更");
    var startRow = 6;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();

    var strFrom = sheet.getRange(1,2).getValue();

    var docBaseID = sheet.getRange(2,2).getValue();
    var strFixedSubject = sheet.getRange(3,2).getValue();

    // テンプレートテキストの取得  
    var docBaseTemplate = DocumentApp.openById(docBaseID);
    var strBaseTemplate = docBaseTemplate.getBody().getText();

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        row.rowNumber = i + startRow;

        // Result列がブランクであれば処理を実行    
        if (!row[10]) { 
            var result = "";

            try
            {
                var strTo = row[0];
                var strCc = row[1];
                var strDestinationSubject = row[2];

                // メールの件名を作成
                var strSubject = "※重要【" + strDestinationSubject + "】" + strFixedSubject;

                var options = {};
                options.cc = strCc;
                options.from = strFrom;

                // 変数を取得
                var strVal1 = row[3];
                var strVal2 = row[4];
                var strVal3 = row[5];
                var strVal4 = row[6];
                var strVal5 = row[7];
                var strVal6 = row[8];
                var strVal7 = row[9];
                
                // メールのbase部分の変数を置換
                var strBody = strBaseTemplate.replace("\{VALUE1\}",strVal1).replace("\{VALUE2\}",strVal2).replace("\{VALUE3\}",strVal3).replace("\{VALUE4\}",strVal4).replace("\{VALUE5\}",strVal5).replace("\{VALUE6\}",strVal6).replace("\{VALUE7\}",strVal7); 

                // メール送信実行       
                GmailApp.sendEmail(strTo,strSubject,strBody,options);

                result = "Success"; 
            }catch(e){
                result = "Error:" + e;
            }

            // 実行結果をResult列にセット
            sheet.getRange(row.rowNumber, lastColum).setValue(result); 
        }
    }  
}






