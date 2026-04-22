// ============================================================
// Craft-Mind 見積書 自動生成スクリプト
// Google Apps Script として このスプレッドシートにバインドして使用
// デプロイ → ウェブアプリ → 全員がアクセス可能
// ============================================================

var TEMPLATE_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
var OUTPUT_FOLDER_NAME = "Craft-Mind 見積書";

function doPost(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json"
  };

  try {
    var data = JSON.parse(e.postData.contents);
    var result = createEstimate(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // CORS preflight 対応
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function createEstimate(data) {
  // 出力フォルダを取得または作成
  var folder = getOrCreateFolder(OUTPUT_FOLDER_NAME);

  // テンプレートシートをコピー
  var templateSS = SpreadsheetApp.openById(TEMPLATE_ID);
  var newSS = templateSS.copy("御見積_" + data.title + "_" + data.estNum);
  folder.addFile(DriveApp.getFileById(newSS.getId()));
  // マイドライブのルートから移動
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(newSS.getId()));

  // ======== 見積書シート（1枚目）を書き込む ========
  var estimateSheet = newSS.getSheets()[0];
  writeEstimateSheet(estimateSheet, data);

  // ======== 利益計算シート（2枚目）を書き込む ========
  if (newSS.getSheets().length > 1) {
    var profitSheet = newSS.getSheets()[1];
    writeProfitSheet(profitSheet, data);
  }

  // スプレッドシートを保存
  SpreadsheetApp.flush();

  // ======== PDF化 ========
  var pdfFile = exportToPdf(newSS, folder, "御見積_" + data.title + "_" + data.estNum + ".pdf");

  // ======== 結果を返す ========
  return {
    status: "success",
    sheetUrl: "https://docs.google.com/spreadsheets/d/" + newSS.getId(),
    pdfUrl: "https://drive.google.com/file/d/" + pdfFile.getId() + "/view",
    pdfId: pdfFile.getId()
  };
}

function writeEstimateSheet(sheet, data) {
  // ---- 宛先（A1セル近辺） ----
  // 既存シートのレイアウト：A1に宛先、件名・納期・支払条件・有効期限が続く
  // 実際のセル位置はシート構造に合わせて調整
  
  var rows = sheet.getDataRange().getValues();
  
  // シートを全スキャンして書き換えるべきセルを特定
  for (var r = 0; r < rows.length; r++) {
    for (var c = 0; c < rows[r].length; c++) {
      var cell = sheet.getRange(r+1, c+1);
      var val = String(rows[r][c]);
      
      // 宛先
      if (val.includes("株式会社ハイテクノ") && c > 0) {
        cell.setValue(data.client);
      }
      // 件名
      if (val === "豊洲ON") {
        cell.setValue(data.title);
      }
      // 納期
      if (val === "2026/3/31" || val === "2026-03-31") {
        if (rows[r][c-1] && String(rows[r][c-1]).includes("納期")) {
          cell.setValue(new Date(data.due));
          cell.setNumberFormat("yyyy/m/d");
        }
      }
      // 有効期限
      if (String(val).includes("2026年3月31日")) {
        cell.setValue(data.validity);
      }
    }
  }

  // ---- 費目の書き込み ----
  // 既存シートの費目行を探す
  var itemStartRow = findItemStartRow(sheet);
  if (itemStartRow > 0) {
    // 既存の費目行をクリア（テンプレートの1行目だけ残して消す）
    // まず既存費目行を探してクリア
    for (var i = 0; i < 20; i++) {
      var checkRow = sheet.getRange(itemStartRow + i, 1, 1, 8).getValues()[0];
      // 費目欄の行かチェック（空でなければクリア）
      if (checkRow[0] !== "" || checkRow[2] !== "") {
        sheet.getRange(itemStartRow + i, 1, 1, 8).clearContent();
      }
    }

    // 費目を書き込む
    for (var idx = 0; idx < data.items.length; idx++) {
      var item = data.items[idx];
      var row = itemStartRow + idx;
      sheet.getRange(row, 1).setValue(item.name);
      sheet.getRange(row, 3).setValue(item.qty);
      sheet.getRange(row, 5).setValue(item.unit);
      sheet.getRange(row, 6).setValue(item.price);
      sheet.getRange(row, 7).setValue(item.qty * item.price);
      if (item.memo) sheet.getRange(row, 8).setValue(item.memo);
    }
  }
}

function writeProfitSheet(sheet, data) {
  // 利益計算シートにも同じく書き込み（仕入れ列含む）
  var rows = sheet.getDataRange().getValues();
  
  for (var r = 0; r < rows.length; r++) {
    for (var c = 0; c < rows[r].length; c++) {
      var val = String(rows[r][c]);
      if (val.includes("株式会社ハイテクノ") && c > 0) {
        sheet.getRange(r+1, c+1).setValue(data.client);
      }
      if (val === "豊洲ON") {
        sheet.getRange(r+1, c+1).setValue(data.title);
      }
      if (String(val).includes("2026年3月31日")) {
        sheet.getRange(r+1, c+1).setValue(data.validity);
      }
    }
  }

  var itemStartRow = findItemStartRow(sheet);
  if (itemStartRow > 0) {
    for (var i = 0; i < 20; i++) {
      sheet.getRange(itemStartRow + i, 1, 1, 8).clearContent();
    }
    for (var idx = 0; idx < data.items.length; idx++) {
      var item = data.items[idx];
      var row = itemStartRow + idx;
      sheet.getRange(row, 1).setValue(item.name);
      sheet.getRange(row, 3).setValue(item.qty);
      sheet.getRange(row, 5).setValue(item.unit);
      sheet.getRange(row, 6).setValue(item.price);
      sheet.getRange(row, 7).setValue(item.qty * item.price);
      if (item.memo) sheet.getRange(row, 8).setValue(item.memo);
      // 仕入れ列（I列 = 9列目）
      if (item.purchase) sheet.getRange(row, 9).setValue(item.purchase);
    }
  }
}

function findItemStartRow(sheet) {
  var rows = sheet.getDataRange().getValues();
  for (var r = 0; r < rows.length; r++) {
    if (String(rows[r][0]) === "作業費") return r + 1;
    // 費目ヘッダー行の次を返す
    if (String(rows[r][0]) === "費目") return r + 2;
  }
  return -1;
}

function exportToPdf(ss, folder, filename) {
  var ssId = ss.getId();
  var sheetId = ss.getSheets()[0].getSheetId();
  
  var url = "https://docs.google.com/spreadsheets/d/" + ssId + 
            "/export?format=pdf" +
            "&size=A4" +
            "&portrait=true" +
            "&fitw=true" +
            "&sheetnames=false" +
            "&printtitle=false" +
            "&pagenumbers=false" +
            "&gridlines=false" +
            "&fzr=false" +
            "&gid=" + sheetId;

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { "Authorization": "Bearer " + token }
  });

  var blob = response.getBlob().setName(filename);
  var pdfFile = folder.createFile(blob);
  return pdfFile;
}

function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

// ============================================================
// テスト用関数（GASエディタで手動実行して確認できる）
// ============================================================
function testCreateEstimate() {
  var testData = {
    estNum: "2604-001",
    client: "テスト株式会社　御中",
    title: "テスト案件",
    due: "2026-05-31",
    validity: "2026年5月21日",
    payment: "貴社ご規定通り",
    items: [
      { name: "作業費", qty: 1, unit: "式", price: 60000, memo: "", purchase: 50000 },
      { name: "部材費", qty: 2, unit: "個", price: 5000, memo: "LANケーブル", purchase: 3000 }
    ]
  };
  
  var result = createEstimate(testData);
  Logger.log(JSON.stringify(result));
}
