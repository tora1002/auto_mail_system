// グローバル変数
var ss = SpreadsheetApp.getActiveSpreadsheet();

// 実行メニューを作成
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu("メールメニュー");
    menu.addItem("領収書不備", "sendReceiptMistakeHtml");
    menu.addItem("収支確認", "sendBalanceCheckHtml");
    menu.addItem("日経テレコン利用ID・PW変更", "sendNikkeiHtml");
    menu.addItem("未着・不備請求書", "sendInvoiceMistakeHtml");
    menu.addItem("検収チェックシート捺印", "sendSealHtml");
    menu.addItem("新規取引外注先", "sendNewSubcontractorHtml");
    menu.addItem("検修書提出", "sendInspectionBookHtml");
    menu.addToUi();
}

function sendReceiptMistakeHtml() {
    var popup = Browser.msgBox("スクリプトを実行しますか？",Browser.Buttons.OK_CANCEL);
    if (popup == "cancel") exit;
    
    var sheet = ss.getSheetByName("領収書不備");
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
        if (!row[11]) { 
            var result = "";

            try
            {
                var strTo = row[0];
                var strCc = row[1];

                // メールの件名を作成
                var strSubject = "【" + accountingMonth + "月経費】" + strFixedSubject + "（" + row[2] + "）";
                
                // メール本文を作成
                var html = "<div> " + row[3] + " さん</div>";
                html += "<br />";
                html += "<div>お疲れ様です。</div>";
                html += "<div>ベクトル管理部の" + row[4] + " です。</div>";
                html += "<br />";
                html += "<div>経費精算にて領収書に不備がありました。</div>";
                html += "<div>下記内容をご確認の上、期日までに領収書の提出をお願いいたします。</div>";
                html += "<br />";
                html += "<div>期日までの提出が難しい場合は連絡ください。</div>";
                html += "<br />";
                html += "<div>また、連絡なく期限までにご提出いただけない場合、</div>";
                html += "<div>翌月の経費精算より相殺させていただきます。</div>";
                html += "<br />";
                html += "<div>提出期限： " + row[5] + "</div>";
                html += "<div>提出先　： " + row[6] + "</div>";
                html += "<br />";
                
                // メールの可変部分を作成
                html += "<div>========================================</div>";
                html += "<br />";
                html += "<div>支払先：　　" + row[7] +  "</div>";
                html += "<div>金額：　　" + row[8] + "</div>";
                html += "<div>状態：　　" + row[9] + "</div>";
                html += "<div>備考：　　" + row[10] + "</div>";

                while (data[i+1] != undefined && strTo == data[i+1][0]) {
                    html += "<br />";
                    html += "<div>支払先：　　" + data[i+1][7] +  "</div>";
                    html += "<div>金額：　　" + data[i+1][8] + "</div>";
                    html += "<div>状態：　　" + data[i+1][9] + "</div>";
                    html += "<div>備考：　　" + data[i+1][10] + "</div>";

                    i = i + 1;
                }

                html += "<br />";
                html += "<div>========================================</div>";

                html += "<br />";
                html += "<div>以上、ご確認のほどよろしくお願いいたします。</div>";

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

function sendBalanceCheckHtml() {
    var popup = Browser.msgBox("スクリプトを実行しますか？",Browser.Buttons.OK_CANCEL);
    if (popup == "cancel") exit;
    
    var sheet = ss.getSheetByName("収支確認");
    var startRow = 5;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();
    var strFrom = sheet.getRange(1,2).getValue();
    var strSubject = sheet.getRange(2,2).getValue();

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
                
                // メール本文を作成
                var html = "<div> " + row[2] + " さん</div>";
                html += "<br />";
                html += "<div>お疲れ様です。</div>";
                html += "<div>ベクトル管理部の" + row[3] + " です。</div>";
                html += "<br />";
                html += "<div>決算のため、下記案件について確認させてください。</div>";
                html += "<br />";

                // メールの可変部分を作成
                html += "<div>========================================</div>";
                html += "<br />";
                html += "<div>JOBNo：　　　" + row[5] +  "</div>";
                html += "<div>案件名：　　" + row[6] + "</div>";
                html += "<div>総利益率：　" + row[7] + "</div>";
                html += "<div>売上計上日：" + row[8] + "</div>";

                while (data[i+1] != undefined && strTo == data[i+1][0]) {
                    html += "<br />";
                    html += "<div>JOBNo：　　　" + data[i+1][5] +  "</div>";
                    html += "<div>案件名：　　" + data[i+1][6] + "</div>";
                    html += "<div>総利益率：　" + data[i+1][7] + "</div>";
                    html += "<div>売上計上日：" + data[i+1][8] + "</div>";

                    i = i + 1;
                }

                html += "<br />";
                html += "<div>========================================</div>";

                html += "<br />";
                html += "<div>利益率が高い案件となっております。</div>";
                html += "<div>監査法人が確認する対象になりますので、</div>";
                html += "<div>下記ご回答いただきますようお願い致します。</div>";
                html += "<br />";
                html += "<div>１. 納品完了日（作業完了日）は上記売上計上日でよろしいでしょうか。</div>";
                html += "<div>２. 追加原価はありませんでしょうか。</div>";
                html += "<div>　追加原価がある場合は、その詳細をご教示ください。</div>";
                html += "<div>　⇒ 支払先、金額等、ZAC登録済みの場合は、JOBNoも併せてお知らせください。</div>";
                html += "<br />";
                html += "<div>　また、追加原価がある場合は、証憑書類(請求書等)を必ず添付のうえ、返信ください。</div>";
                html += "<br />";
                html += "<div>　(1) 納品日：</div>";
                html += "<div>　(2) 追加原価</div>";
                html += "<div>　　･支払先：</div>";
                html += "<div>　　･金額：</div>";
                html += "<div>　　･（zac登録済みの場合）JOBNo：</div>";
                html += "<br />";
                html += "<div>回答期日： " + row[4] + " まで</div>";
                html += "<br />";
                html += "<div>お忙しいところ大変恐れ入りますが、</div>";
                html += "<div>ご確認、ご返答の程宜しくお願い致します。</div>";

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

function sendNikkeiHtml() {
    var popup = Browser.msgBox("スクリプトを実行しますか？",Browser.Buttons.OK_CANCEL);
    if (popup == "cancel") exit;
    
    var sheet = ss.getSheetByName("日経テレコン利用ID・PW変更");
    var startRow = 5;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();
    var strFrom = sheet.getRange(1,2).getValue();
    var strFixedSubject = sheet.getRange(2,2).getValue();

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        row.rowNumber = i + startRow;

        // Result列がブランクであれば処理を実行    
        if (!row[7]) { 
            var result = "";

            try
            {
                var strTo = row[0];
                var strCc = row[1];

                // メールの件名を作成
                var strSubject = "※重要【" + row[2] + "】" + strFixedSubject;
                
                // メールの本文を作成
                var html = "<div>" + row[2] + " 各位</div>";
                html += "<br />";
                html += "<div>お疲れ様です。</div>";
                html += "<div>ベクトル管理部の" + row[3] + "です。</div>";
                html += "<br />";
                html += "<div>新事業年度部署編成に伴い、日経テレコン利用IDを各部署振り直しました。</div>";
                html += "<div>本日（" + row[4] + "）より下記ID・パスワードにて日経テレコンをご利用ください。</div>";
                html += "<div>※IDに変更のなかったチームもパスワードは変更しております。</div>";
                html += "<br />";
                html += "<div>" + row[2] + "</div>";
                html += "<div>ID： " + row[5] + "</div>";
                html += "<div>PW： " + row[6] + "</div>";
                html += "<br />";
                html += "<div>何かございましたら " + row[3] + " までお問い合わせください。</div>";
                html += "<div>以上、ご確認のほどよろしくお願いいたします。</div>";

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

function sendInvoiceMistakeHtml() {
    var popup = Browser.msgBox("スクリプトを実行しますか？",Browser.Buttons.OK_CANCEL);
    if (popup == "cancel") exit;
    
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

function sendSealHtml() {
    var popup = Browser.msgBox("スクリプトを実行しますか？",Browser.Buttons.OK_CANCEL);
    if (popup == "cancel") exit;
    
    var sheet = ss.getSheetByName("検収チェックシート捺印");
    var startRow = 5;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();
    var strFrom = sheet.getRange(1,2).getValue();
    var strSubject = sheet.getRange(2,2).getValue();

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
                var strVal1 = row[2];
                var strVal2 = row[3];
                var strVal3 = row[4];
                
                // メールの本文を作成
                var html = "<div>" + strVal1 + " さん</div>";
                html += "<br />";
                html += "<div>お疲れ様です。</div>";
                html += strVal2 + " 月収支表のご提出ありがとうございます。</div>";
                html += "<br />";
                html += "<div>検収チェックシートに上長印の捺印を頂きたく、</div>";
                html += "<div>収支表をデスクの上に置かせていただきました。</div>";
                html += "<br />";
                html += "<div>" + strVal3 + " が期日となっておりますので、</div>";
                html += "<div>お手数ではございますが、</div>";
                html += "<div>ご対応いただきましたら収支表を管理部にお戻し下さい。</div>";
                html += "<br />";
                html += "<div>何卒よろしくお願いいたします。</div>";

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

function sendNewSubcontractorHtml() {
    var popup = Browser.msgBox("スクリプトを実行しますか？",Browser.Buttons.OK_CANCEL);
    if (popup == "cancel") exit;
    
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

function sendInspectionBookHtml() {
    var popup = Browser.msgBox("スクリプトを実行しますか？",Browser.Buttons.OK_CANCEL);
    if (popup == "cancel") exit;
    
    var sheet = ss.getSheetByName("検修書提出");
    var startRow = 5;
    
    var lastColum = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataRange = sheet.getRange(startRow, 1, numRows, lastColum);
    var data = dataRange.getValues();

    var strFrom = sheet.getRange(1,2).getValue();
    var strFixedSubject = sheet.getRange(2,2).getValue();

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        row.rowNumber = i + startRow;

        // Result列がブランクであれば処理を実行    
        if (!row[17]) { 
            var result = "";

            try
            {
                var strTo = row[0];
                var strCc = row[1];

                // メールの件名を作成
                var strSubject = "※重要※【" + row[2] + "】" + strFixedSubject + " / " + row[3]  + "回目連絡";

                // メールの本文を作成
                var html = "<div>ご担当者様</div>";
                html += "<br />";
                html += "<div>お疲れ様です。</div>";
                html += "<div>ベクトル管理部の " + row[4] + " です。</div>";
                html += "<br />";
                html += "<div>回収状況を更新しましたのでご確認をお願いいたします。</div>";
                html += "<br />";

                // 表の見出し部分を作成
                html += "<table style='border-collapse:collapse;'>";
                html += "<tr>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;' bgcolor='#ffffc0'>検収書チェック</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>JOBNo</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>受託会社</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>案件名</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>営業担当者</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>担当部門</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>請求先名</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>売上予定日</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>売上日</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>売上計上(予定実績)額</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>プロジェクトコード(メイン)</th>";
                html += "<th style='border:1px solid #ccc; padding:3px 5px;'>プロジェクト名(メイン)</th>";
                html += "</tr>";
                
                // 表のデータ部分を作成
                html += "<tr>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;' bgcolor='#ffffc0'>" + row[5] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[6] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[7] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[8] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[9] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[10] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[11] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[12] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[13] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[14] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[15] + "</td>";
                html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + row[16] + "</td>";
                html += "</tr>";
                
                while (data[i+1] != undefined && strTo == data[i+1][0]) {
                    html += "<tr>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;' bgcolor='#ffffc0'>" + data[i+1][5] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][6] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][7] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][8] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][9] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][10] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][11] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][12] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][13] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][14] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][15] + "</td>";
                    html += "<td style='border:1px solid #ccc; padding:3px 5px;'>" + data[i+1][16] + "</td>";
                    html += "</tr>";

                    i = i + 1;
                }
                
                html += "</table>";

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


