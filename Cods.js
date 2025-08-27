function doGet(e) {
  const t = HtmlService.createTemplateFromFile('form');
  t.isLocal = false; // 本番なら false, ローカル確認なら true
  return t.evaluate();
}

// マスターデータを取得する関数
function getMasterData() {
  const ss = SpreadsheetApp.openById('1mXMe5UFKDVPpB4VSens3s_NBYMrfA-e6zbaN3gSBuZA');


  const companySheet = ss.getSheetByName('元請会社マスタ');
  const staffSheet = ss.getSheetByName('TTC担当者名マスタ');
  const siteSheet = ss.getSheetByName('現場名マスタ');
  const workerSheet = ss.getSheetByName('作業者名マスタ');
  const partnerSheet = ss.getSheetByName('協力会社名マスタ');


  const company = companySheet.getRange(2, 1, companySheet.getLastRow() - 1, 2).getValues()
    .map(([id, name]) => ({ id, name }));


  const staff = staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, 3).getValues()
    .map(([id, name, companyIds]) => {
      const companies = companyIds.split(',').map(c => c.trim());
      return { id, name, companies };
    });


  const site = siteSheet.getRange(2, 1, siteSheet.getLastRow() - 1, 4).getValues()
    .map(([id, name, staffId, companyIds]) => {
      const companies = (companyIds || '').split(',').map(c => c.trim());
      return { id, name, staffId, companyIds: companies };
    });




const worker = workerSheet.getRange(2, 1, workerSheet.getLastRow() - 1, 3).getValues()
  .map(([id, name, staffIds]) => {
    const staffs = (staffIds || '').split(',').map(s => s.trim());
    return { id, name, staffIds: staffs };
  });


const partner = partnerSheet.getRange(2, 1, partnerSheet.getLastRow() - 1, 3).getValues()
  .map(([id, name, staffIds]) => {
    const staffs = (staffIds || '').split(',').map(s => s.trim());
    return { id, name, staffIds: staffs };
  });


  return { company, staff, site, worker, partner };
  Logger.log(JSON.stringify(result));
  return result;
}


function isValidQuarterHour(value) {
  return Number.isFinite(value) && Math.round(value * 100) % 25 === 0;
}


// 回答を保存する関数
function saveFormResponse(data) {
  const ss = SpreadsheetApp.openById('1mXMe5UFKDVPpB4VSens3s_NBYMrfA-e6zbaN3gSBuZA');
  const sheet = ss.getSheetByName('回答');
  const timestamp = new Date();
  const reportDate = data.date || Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");


  // マスタ取得
  const companyMap = getMapFromSheet(ss.getSheetByName('元請会社マスタ'));
  const siteMap = getMapFromSheet(ss.getSheetByName('現場名マスタ'));
  const staffMap = getMapFromSheet(ss.getSheetByName('TTC担当者名マスタ'));


  const companyName = companyMap[data.companyId] || data.companyId;
  const siteName = siteMap[data.site] || data.site;
  const staffName = staffMap[data.staffId] || data.staffId;


  // 作業者
if (Array.isArray(data.workers)) {
  data.workers.forEach(w => {
    sheet.appendRow([
      timestamp, reportDate,
      companyName, staffName, siteName,
      '作業者',
      w.name,
//      toNumber(w.day), toNumber(w.evening), toNumber(w.night), toNumber(w.overtime)
      toNumber(w.man), toNumber(w.overtime)
    ]);
  });
}

// 協力会社
if (Array.isArray(data.partners)) {
  data.partners.forEach(p => {
    sheet.appendRow([
      timestamp, reportDate,
      companyName, staffName, siteName,
      '協力会社',
      p.name,
//      toNumber(p.day), toNumber(p.evening), toNumber(p.night), toNumber(p.overtime)
      toNumber(p.man), toNumber(p.overtime)
    ]);
  });
}

  Logger.log(data);
  return 'ok';
}


/**
 * A列：ID、B列：名前 のマスタを連想配列にする
 */
function getMapFromSheet(sheet) {
  const values = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < values.length; i++) { // 1行目はヘッダー想定
    const id = values[i][0];
    const name = values[i][1];
    if (id && name) {
      map[id] = name;
    }
  }
  return map;
}


// 過去の回答を取得する関数
function getPreviousRecords(date, companyId, staffId, siteId) {
  const ss = SpreadsheetApp.openById('1mXMe5UFKDVPpB4VSens3s_NBYMrfA-e6zbaN3gSBuZA');
  const sheet = ss.getSheetByName('回答');
  const values = sheet.getDataRange().getValues();

  const companyMap = getMapFromSheet(ss.getSheetByName('元請会社マスタ'));
  const siteMap = getMapFromSheet(ss.getSheetByName('現場名マスタ'));
  const staffMap = getMapFromSheet(ss.getSheetByName('TTC担当者名マスタ'));

  const companyName = companyMap[companyId];
  const siteName = siteMap[siteId];
  const staffName = staffMap[staffId];

  const filtered = values.filter(row =>
    Utilities.formatDate(new Date(row[1]), "Asia/Tokyo", "yyyy-MM-dd") === date &&
    row[2] === companyName &&
    row[3] === staffName &&
    row[4] === siteName
  ).map(row => ({
    type: row[5],
    name: row[6],
    // day: row[7],
    // evening: row[8],
    man: row[7],
    overtime: row[8],
  }));

  return filtered;
}


// 以下の関数は、Google Sheetsを使用してデータを取得、保存、および更新するためのものです。
function updateEditedRecords(meta, records) {
  const ss = SpreadsheetApp.openById('1mXMe5UFKDVPpB4VSens3s_NBYMrfA-e6zbaN3gSBuZA');
  const sheet = ss.getSheetByName('回答');
  const logSheet = ss.getSheetByName('編集ログ') || ss.insertSheet('編集ログ');

  const companyMap = getMapFromSheet(ss.getSheetByName('元請会社マスタ'));
  const siteMap = getMapFromSheet(ss.getSheetByName('現場名マスタ'));
  const staffMap = getMapFromSheet(ss.getSheetByName('TTC担当者名マスタ'));

  const date = meta.date;
  const companyName = companyMap[meta.companyId];
  const siteName = siteMap[meta.siteId];
  const staffName = staffMap[meta.staffId];

  const all = sheet.getDataRange().getValues();
  const newTimestamp = new Date();

  // 全行チェックして、条件に一致した行に処理
  for (let i = all.length - 1; i >= 1; i--) {
    const r = all[i];
    if (
      Utilities.formatDate(new Date(r[1]), "Asia/Tokyo", "yyyy-MM-dd") === date &&
      r[2] === companyName &&
      r[3] === staffName &&
      r[4] === siteName
    ) {
      const rowIndex = i + 1; // シート上の行番号（1スタート）

      // 編集ログへの追記を試みる
      let logSuccess = false;  // 初期値のフラグ→最初はfalseにして成功したらフラグが立つようにする
      try {
        logSheet.appendRow(["編集前", ...r]);
        logSuccess = true;  // ここで成功フラグを立てる→もし成功してない場合には検知してリトライ
      } catch (e) {
        // 1回だけリトライ
        try {
          Utilities.sleep(1000); // 少し待機
          logSheet.appendRow(["編集前", ...r]);
          logSuccess = true;
        } catch (retryError) {
          // リトライも失敗 → L列に「編集ログへの転記エラー」
          sheet.getRange(rowIndex, 12).setValue("編集ログへの転記エラー");
        }
      }

      // ログ追記が成功しているときのみ削除を試みる
      if (logSuccess) {
        try {
          sheet.deleteRow(rowIndex);  // 行を削除
        } catch (e) {
          // 削除失敗した場合もエラー記録
          sheet.getRange(rowIndex, 12).setValue("編集ログへの転記エラー");
        }
      }
    }
  }

  // 編集後データの追加処理
  records.forEach(r => {
    sheet.appendRow([
      newTimestamp, date,
      companyName, staffName, siteName,
      r.type,
      r.name,
      toNumber(r.man), toNumber(r.overtime)
    ]);
  });

  return 'ok';
}


function toNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  const n = parseFloat(val);
  return Number.isFinite(n) ? n : 0;
}

