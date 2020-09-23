var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];

// 実行メニューを作成
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu("GAS実行");
    menu.addItem("メール送信実行", "sendMergeEmail");
    menu.addToUi();
}

function sendMergeEmail(){
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var startRow = 6;
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();

    var strFrom = sheet.getRange(1,2).getValue();

    var docID = sheet.getRange(2,2).getValue();
    var attachementID = sheet.getRange(3,2).getValue();

    // テンプレートテキストの取得  
    var docTemplate = DocumentApp.openById(docID);
    var strTemplate = docTemplate.getBody().getText();

    for (var i = 0; i < data.length; ++i) {
        var row = data[i];
        row.rowNumber = i + startRow;

        // Result列がブランクであれば処理を実行    
        if (!row[7]) { 
            var result = "";

            try
            {
                var strVal1 = row[4];
                var strVal2 = row[5];
                var strVal3 = row[6];
                var strVal4 = row[7];
                var strVal5 = row[8];
                var strVal6 = row[9];
                var strVal7 = row[10];

                // テンプレートテキスト内の変数を置換
                var strBody = strTemplate.replace("\{VALUE1\}",strVal1).replace("\{VALUE2\}",strVal2).replace("\{VALUE3\}",strVal3).replace("\{VALUE4\}",strVal4).replace("\{VALUE5\}",strVal5).replace("\{VALUE6\}",strVal6).replace("\{VALUE7\}",strVal7); 

                var strTo = row[0];
                var strCc = row[1];
                var strBcc = row[2];
                var strSubject = row[3];

                var options = {};
                options.cc = strCc;
                options.bcc = strBcc;
                options.from = strFrom;

                // 添付ファイル指定がある場合はoptionsに追加（※未使用）
                //if(attachementID){
                //    var attachment = DriveApp.getFileById(attachementID);
                //    options.attachments = attachment
                //}

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

