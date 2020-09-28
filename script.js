// グローバル変数
var ss = SpreadsheetApp.getActiveSpreadsheet();

// 実行メニューを作成
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu("メールメニュー");
    menu.addItem("領収書不備", "sendReceiptMistake");
    menu.addItem("収支確認", "sendBalanceCheck");
    menu.addItem("日経テレコン利用ID・PW変更", "sendNikkei");
    //menu.addItem("未着・不備請求書", "sendInvoiceMistake");
    menu.addItem("未着・不備請求書", "sendInvoiceMistakeHtml");
    menu.addItem("検収チェックシート捺印", "sendSeal");
    //menu.addItem("新規取引外注先", "sendNewSubcontractor");
    menu.addItem("新規取引外注先", "sendNewSubcontractorHtml");
    //menu.addItem("検修書提出", "sendInspectionBook");
    menu.addToUi();
}

function sendReceiptMistake() {
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

function sendBalanceCheck() {
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

function sendNikkei() {
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
                
                // メールの変数を置換
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

function sendInvoiceMistake() {
    var sheet = ss.getSheetByName("未着・不備請求書");
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
                
                // 使うか確認？
                var strDestinationSubject = row[2];

                // メールの件名を作成
                var strSubject = strFixedSubject + "（" + accountingMonth + "月分）";

                var options = {};
                options.cc = strCc;
                options.from = strFrom;

                // メールの変数を取得
                var strVal1 = row[3];
                var strVal2 = row[4];
                
                // メールのbase部分の変数を置換
                var strBody = strBaseTemplate.replace("\{VALUE1\}",strVal1).replace("\{VALUE2\}",strVal2); 

                // メールの表のヘッダーを作成
                var strVariable = "請求書日付　支払先　金額　内容　ステータス\n";

                // メールの表の可変部分の変数を取得
                var strVal4 = row[5];
                var strVal5 = row[6];
                var strVal6 = row[7];
                var strVal7 = row[8];
                var strVal8 = row[9];

                var strVariable = strVariable + strVal4 + "　" + strVal5 + "　" + strVal6 + "　" + strVal7 + "　" + strVal8 + "\n";

                while (data[i+1] != undefined && strTo == data[i+1][0]) {
                    var strVal4 = data[i+1][5];
                    var strVal5 = data[i+1][6];
                    var strVal6 = data[i+1][7];
                    var strVal7 = data[i+1][8];
                    var strVal8 = data[i+1][9];

                    var strVariable = strVariable + strVal4 + "　" + strVal5 + "　" + strVal6 + "　" + strVal7 + "　" + strVal8 + "\n";

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

function sendInvoiceMistakeHtml() {
    var sheet = ss.getSheetByName("未着・不備請求書");
    var startRow = 6;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();

    var strFrom = sheet.getRange(1,2).getValue();
    var accountingMonth = sheet.getRange(2,2).getValue();
    var strFixedSubject = sheet.getRange(3,2).getValue();

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
                var strSubject = "【" + strDestinationSubject  + "】" + strFixedSubject + "（" + accountingMonth + "月分）";

                // メールの変数を取得
                var strVal1 = row[3];
                var strVal2 = row[4];

                // メールの本文を作成
                var html = "<div>ご担当者様</div>";
                html += "<br />";
                html += "<div>お疲れ様です。</div>";
                html += "<div>収支表のご提出ありがとうございました。</div>";
                html += "<br />";
                html += "<div>下記の請求書等に不備がございます。</div>";
                html += "<div>手配していただき、提出期日までにご提出をお願いいたします。</div>";
                html += "<br />";
                html += "<div style='font-weight: bold; text-decoration: underline; color: #FF0000;'>提出期日： " + strVal1 + "</div>";
                html += "<div style='font-weight: bold;'>※ご提出の際は " + strVal2 + " さんまでお願い致します。</div>";
                html += "<br />";

                // 表の見出し部分を作成
                html += "<table style='border-collapse:collapse;'>";
                html += "<tr bgcolor='#ffffc0'>";
                html += "<th style='border:1px solid #ccc; padding:10px;'>請求書日付</th>";
                html += "<th style='border:1px solid #ccc; padding:10px;'>支払先</th>";
                html += "<th style='border:1px solid #ccc; padding:10px;'>金額</th>";
                html += "<th style='border:1px solid #ccc; padding:10px;'>内容</th>";
                html += "<th style='border:1px solid #ccc; padding:10px;'>ステータス</th>";
                html += "</tr>";

                // 表のデータ部分を作成
                html += "<tr>";
                html += "<td style='border:1px solid #ccc; padding:10px;'>" + row[5] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:10px;'>" + row[6] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:10px;'>" + row[7] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:10px;'>" + row[8] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:10px;'>" + row[9] + "</td>";
                html += "</tr>";

                while (data[i+1] != undefined && strTo == data[i+1][0]) {
                    html += "<tr>";
                    html += "<td style='border:1px solid #ccc; padding:10px;'>" + data[i+1][5] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:10px;'>" + data[i+1][6] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:10px;'>" + data[i+1][7] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:10px;'>" + data[i+1][8] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:10px;'>" + data[i+1][9] + "</td>";
                    html += "</tr>";

                    i = i + 1;
                }

                html += "</table>";
                html += "<br />";

                html += "<div>----------------------------------------</div>";
                html += "<div style='font-weight: bold; color: #FF0000;'>【請求書なしの請求書とは・・・】</div>";
                html += "<div style='font-weight: bold;'>■　請求書の添付がない</div>";
                html += "<div style='font-weight: bold;'>■　見積書、納品書の添付</div>";
                html += "<br />";
                html += "<div style='font-weight: bold; color: #FF0000;'>【原本なしの請求書とは・・・】</div>";
                html += "<div style='font-weight: bold;'>■　金額違い</div>";
                html += "<div style='font-weight: bold;'>■　社判なし</div>";
                html += "<div style='font-weight: bold;'>■　PDF請求書</div>";
                html += "<br />";
                html += "<div>PDF請求書のみ発行している企業、個人の場合は、</div>";
                html += "<div style='font-weight: bold; color: #FF0000;'>「原本」と記載して担当者の印鑑</p> を押してください。</div>";
                html += "<div style='font-weight: bold; text-decoration: underline; background-color: #FFFF00;'>※ない場合は原本未着扱いとしてお支払いを致しません。</div>";
                html += "<br />";
                html += "<div>【宛名間違いの請求書とは・・・】</div>";
                html += "<div>例えば、AN,PL,INで受注している案件に</div>";
                html += "<div>ベクトル宛の請求書が発行されている場合が上記にあたります。</div>";
                html += "<div>----------------------------------------</div>";

                var options = {};
                options.cc = strCc;
                options.from = strFrom;
                options.htmlBody = html;

                // メール送信実行       
                GmailApp.sendEmail(strTo, strSubject, "", options);

                result = "Success"; 
            }catch(e){
                result = "Error:" + e;
            }

            // 実行結果をResult列にセット
            sheet.getRange(row.rowNumber, lastColum).setValue(result); 
        }
    }  
}

function sendSeal() {
    var sheet = ss.getSheetByName("検収チェックシート捺印");
    var startRow = 6;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();

    var strFrom = sheet.getRange(1,2).getValue();

    var docBaseID = sheet.getRange(2,2).getValue();
    var strSubject = sheet.getRange(3,2).getValue();

    // テンプレートテキストの取得  
    var docBaseTemplate = DocumentApp.openById(docBaseID);
    var strBaseTemplate = docBaseTemplate.getBody().getText();

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        row.rowNumber = i + startRow;

        // Result列がブランクであれば処理を実行    
        if (!row[5]) { 
            var result = "";

            try
            {
                var strTo = row[0];
                var strCc = row[1];

                var options = {};
                options.cc = strCc;
                options.from = strFrom;

                // 変数を取得
                var strVal1 = row[2];
                var strVal2 = row[3];
                var strVal3 = row[4];
                
                // メールの変数を置換
                var strBody = strBaseTemplate.replace("\{VALUE1\}",strVal1).replace("\{VALUE2\}",strVal2).replace("\{VALUE3\}",strVal3); 

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

function sendNewSubcontractorHtml() {
    var sheet = ss.getSheetByName("新規取引外注先");
    var startRow = 6;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();

    var strFrom = sheet.getRange(1,2).getValue();

    var strVal1 = sheet.getRange(2,2).getValue();
    var strFixedSubject = sheet.getRange(3,2).getValue();

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
                var strSubject = "【" + strDestinationSubject + "】" + strFixedSubject;

                // メールの本文を作成
                var html = "<div>ご担当者様</div>";
                html += "<br />";
                html += "<div>お疲れ様です。</div>";
                html += "<div>" + strVal1 + "月に新規取引が開始された外注先をお送りします。</div>";
                html += "<br />";

                // 表の見出し部分を作成
                html += "<table style='border-collapse:collapse;'>";
                html += "<tr bgcolor='#ffffc0'>";
                html += "<th style='border:1px solid #ccc; padding:10px;'>個人 or 法人</th>";
                html += "<th style='border:1px solid #ccc; padding:10px;'>正式名称</th>";
                html += "<th style='border:1px solid #ccc; padding:10px;'>最終更新者(営業)</th>";
                html += "</tr>";
                
                // 表のデータ部分を作成
                html += "<tr>";
                html += "<td style='border:1px solid #ccc; padding:10px;'>" + row[3] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:10px;'>" + row[4] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:10px;'>" + row[5] + "</td>";
                html += "</tr>";
                
                while (data[i+1] != undefined && strTo == data[i+1][0]) {
                    html += "<tr>";
                    html += "<td style='border:1px solid #ccc; padding:10px;'>" + data[i+1][3] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:10px;'>" + data[i+1][4] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:10px;'>" + data[i+1][5] + "</td>";
                    html += "</tr>";

                    i = i + 1;
                }
                
                html += "</table>";

                html += "<br />";
                html += "<div>「取引登録申請書」（法人or個人）をお送りして</div>";
                html += "<div>記入・捺印をいただいたうえで管理部にご提出ください。</div>";
                html += "<div>管理部宛に直接お送りいただいても構いません。 </div>";
                html += "<br />";
                html += "<div>詳細については <a href='https://sites.google.com/a/vectorinc.co.jp/pp/home/contract/law-rule'>社内ポータル</a> をご確認ください</div>";
                html += "<br />";
                html += "<div>以上、よろしくお願いいたします。</div>";

                var options = {};
                options.cc = strCc;
                options.from = strFrom;
                options.htmlBody = html;

                // メール送信実行       
                GmailApp.sendEmail(strTo, strSubject, "", options);

                result = "Success"; 
            }catch(e){
                result = "Error:" + e;
            }

            // 実行結果をResult列にセット
            sheet.getRange(row.rowNumber, lastColum).setValue(result); 
        }
    }  
}




// 未使用
function sendNewSubcontractor() {
    var sheet = ss.getSheetByName("新規取引外注先");
    var startRow = 7;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();

    var strFrom = sheet.getRange(1,2).getValue();

    var docBaseID = sheet.getRange(2,2).getValue();
    var docVariableID = sheet.getRange(3,2).getValue();

    var strVal1 = sheet.getRange(4,2).getValue();
    var strFixedSubject = sheet.getRange(5,2).getValue();

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
                var strSubject = "【" + strDestinationSubject + "】" + strFixedSubject;

                var options = {};
                options.cc = strCc;
                options.from = strFrom;

                // メールの変数を置換
                var strBody = strBaseTemplate.replace("\{VALUE1\}",strVal1); 

                // メールの表のヘッダーを作成
                var strVariable = "個人or法人　提出先　正式名称\n";

                // メールの表の可変部分の変数を取得
                var strVal2 = row[3];
                var strVal3 = row[4];
                var strVal4 = row[5];

                var strVariable = strVariable + strVal2 + "　" + strVal3 + "　" + strVal4 + "\n";

                while (data[i+1] != undefined && strTo == data[i+1][0]) {
                    var strVal2 = data[i+1][3];
                    var strVal3 = data[i+1][4];
                    var strVal4 = data[i+1][5];

                    var strVariable = strVariable + strVal2 + "　" + strVal3 + "　" + strVal4 + "\n";

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

