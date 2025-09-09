// webアプリを実施する際に一番最初に呼ばれる関数
// htmlをビルドしていく感じ→ form.html を呼び出す
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('form');
  t.isLocal = false; // 本番なら false, ローカル確認なら true
  return t.evaluate();
}

/** =========================
 * 集中管理：スプレッドシート
 * ========================= */
const SHEET_ID = '1R1QHdj1ZVtXwInZqvHBzX57Pumm70wvZDLO2hYt1bpM'; //
const SHEETS = {
  RESPONSES: '回答',
  COMPANY: '元請会社マスタ',
  STAFF: 'TTC担当者名マスタ',
  SITE: '現場名マスタ',
  WORKER: '作業者名マスタ',
  PARTNER: '協力会社名マスタ',
};

function getSS() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// cssファイルのhtmlファイルを呼び出し紐づける関数
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// マスターデータを取得する関数
function getMasterData() {
  const ss = getSS();

  const companySheet = ss.getSheetByName(SHEETS.COMPANY);
  const staffSheet = ss.getSheetByName(SHEETS.STAFF);
  const siteSheet = ss.getSheetByName(SHEETS.SITE);
  const workerSheet = ss.getSheetByName(SHEETS.WORKER);
  const partnerSheet = ss.getSheetByName(SHEETS.PARTNER);


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
  const ss = getSS();
  const sheet = ss.getSheetByName(SHEETS.RESPONSES);
  const timestamp  = new Date();
  const reportDate = data.date || Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");

  // マスタ取得
  const companyMap = getMapFromSheet(ss.getSheetByName('元請会社マスタ'));
  const siteMap    = getMapFromSheet(ss.getSheetByName('現場名マスタ'));
  const staffMap   = getMapFromSheet(ss.getSheetByName('TTC担当者名マスタ'));

  const companyName = companyMap[data.companyId] || data.companyId;
  const siteName    = siteMap[data.site]        || data.site;
  const staffName   = staffMap[data.staffId]    || data.staffId;

  // 役割 → サフィックス（末尾に足す文字列）
  // 「一般作業員」は会社名のみ（サフィックスなし）
  const roleSuffix = {
    only:   '',        // 一般作業員 → 「会社名」
    leader: ' 職長',   // 職長       → 「会社名 職長」
    other:  ' 有資格者'// 有資格者   → 「会社名 有資格者」
  };

  // 作業者
  if (Array.isArray(data.workers)) {
    data.workers.forEach(w => {
      sheet.appendRow([
        timestamp, reportDate,
        companyName, staffName, siteName,
        '作業者',
        w.name,
        toNumber(w.man), toNumber(w.overtime)
      ]);
    });
  }

  // 協力会社（ここで表示名を組み立て）
  if (Array.isArray(data.partners)) {
    data.partners.forEach(p => {
      // p.name は会社名、p.role は 'only' | 'leader' | 'other'（無い場合は空扱い）
      const suffix = roleSuffix[p.role] || '';
      const displayName = (p.name || '') + suffix;

      sheet.appendRow([
        timestamp, reportDate,
        companyName, staffName, siteName,
        '協力会社',
        displayName,
        toNumber(p.man), toNumber(p.overtime)
      ]);
    });
  }

  Logger.log('[saveFormResponse] data=%s', JSON.stringify(data));
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
  Logger.log('--- getPreviousRecords 開始 ---');
  Logger.log('指定日付: %s, 元請ID: %s, 担当ID: %s, 現場ID: %s', date, companyId, staffId, siteId);

  const ss = getSS();
  const sheet = ss.getSheetByName(SHEETS.RESPONSES);
  Logger.log('読み取り中スプシURL: %s', ss.getUrl());

  const values = sheet.getDataRange().getValues();
  Logger.log('全行数=%s 行', values.length);

  const companyMap = getMapFromSheet(ss.getSheetByName('元請会社マスタ'));
  const siteMap = getMapFromSheet(ss.getSheetByName('現場名マスタ'));
  const staffMap = getMapFromSheet(ss.getSheetByName('TTC担当者名マスタ'));

  const companyName = companyMap[companyId];
  const siteName = siteMap[siteId];
  const staffName = staffMap[staffId];

  Logger.log('会社名: %s, 現場名: %s, 担当者名: %s', companyName, siteName, staffName);

  const filtered = values
    .filter(row => {
      // row[1] は「出面日付」列
      const cell = row[1];
      let cellDateStr;

      if (cell instanceof Date) {
        // 日付型ならフォーマット
        cellDateStr = Utilities.formatDate(cell, "Asia/Tokyo", "yyyy-MM-dd");
      } else {
        // 文字列（"yyyy-MM-dd" で入っている想定）
        cellDateStr = String(cell);
      }

      // 🔽 ここで1行ごとにログ出力
      Logger.log(
        '[行の内容] 入力日付=%s, 出面日付=%s | 業種=%s vs %s | 担当者=%s vs %s | 現場名=%s vs %s',
        cellDateStr, date, row[2], companyName, row[3], staffName, row[4], siteName
      );

      return (
        cellDateStr === date &&
        row[2] === companyName &&
        row[3] === staffName  &&
        row[4] === siteName
      );
    })
    .map(row => ({
      type: row[5],
      name: row[6],
      man: row[7],
      overtime: row[8],
    }));

  Logger.log('[getPreviousRecords] 該当行数=%s 件', filtered.length);
  return filtered;
}


// 以下の関数は、Google Sheetsを使用してデータを取得、保存、および更新するためのものです。
function updateEditedRecords(meta, records) {
  const FN = '[updateEditedRecords]';
  const start = new Date();

  // --- utils ---
  const sleep = (ms) => Utilities.sleep(ms);
  const doWithRetry = (fn, { retries = 1, waitMs = 500 } = {}) => {
    let lastErr;
    for (let i = 0; i <= retries; i++) {
      try { return fn(); }
      catch (e) {
        lastErr = e;
        if (i < retries) sleep(waitMs);
      }
    }
    throw lastErr;
  };

  // --- lock ---
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    throw new Error('他の処理が実行中です。しばらくしてから再度お試しください。');
  }

  try {
    // --- resolve masters ---
    const ss     = getSS();
    const sheet  = ss.getSheetByName(SHEETS.RESPONSES);
    const logSh  = ss.getSheetByName('編集ログ') || ss.insertSheet('編集ログ');

    const companyMap = getMapFromSheet(ss.getSheetByName(SHEETS.COMPANY));
    const siteMap    = getMapFromSheet(ss.getSheetByName(SHEETS.SITE));
    const staffMap   = getMapFromSheet(ss.getSheetByName(SHEETS.STAFF));

    const dateStr     = String(meta.date || ''); // yyyy-MM-dd
    const companyName = companyMap[meta.companyId];
    const siteName    = siteMap[meta.siteId];
    const staffName   = staffMap[meta.staffId];

    if (!companyName || !siteName || !staffName) {
      throw new Error(
        'マスター解決に失敗しました（company/site/staffの名前が取得できません）。' +
        JSON.stringify({ companyId: meta.companyId, siteId: meta.siteId, staffId: meta.staffId })
      );
    }

    Logger.log('%s meta=%s', FN, JSON.stringify(meta));
    Logger.log('%s records.length=%s', FN, (records || []).length);
    Logger.log('%s records(raw)=%s', FN, JSON.stringify(records || []));

    const all = sheet.getDataRange().getValues(); // [0] header
    const newTimestamp = new Date();

    // --- find matched rows ---
    const matchRowIndexes = []; // 1-based
    for (let i = 1; i < all.length; i++) {
      const r = all[i];

      // r[1] may be Date or string
      let rowDateStr = '';
      try {
        const cell = r[1];
        const d = (cell instanceof Date) ? cell : (cell ? new Date(cell) : null);
        if (d && !isNaN(d.getTime())) {
          rowDateStr = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
        } else {
          rowDateStr = String(cell || '');
        }
      } catch (e) {}

      if (rowDateStr === dateStr && r[2] === companyName && r[3] === staffName && r[4] === siteName) {
        matchRowIndexes.push(i + 1); // to 1-based
      }
    }
    Logger.log('%s matched rows=%s', FN, JSON.stringify(matchRowIndexes));

    // --- log & delete matched rows (bottom-up) ---
    matchRowIndexes.sort((a, b) => b - a).forEach(rowIndex => {
      const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

      let logged = false;
      try {
        doWithRetry(() => logSh.appendRow(['編集前', ...row]), { retries: 1, waitMs: 1000 });
        logged = true;
      } catch (e) {
        Logger.log('%s log append failed row=%s err=%s', FN, rowIndex, e);
        try { sheet.getRange(rowIndex, 12).setValue('編集ログへの転記エラー'); } catch (e2) {}
      }

      if (logged) {
        try {
          doWithRetry(() => sheet.deleteRow(rowIndex), { retries: 1, waitMs: 600 });
        } catch (e) {
          Logger.log('%s delete failed row=%s err=%s', FN, rowIndex, e);
          try { sheet.getRange(rowIndex, 12).setValue('行削除エラー'); } catch (e2) {}
        }
      }
    });

    SpreadsheetApp.flush();

    // --- append new rows (compose name with role for 協力会社) ---
const roleSuffixMap = {
  only:  '',
  leader:' 職長',
  other: ' 有資格者'
};

(records || []).forEach(r => {
  const nameForSave =
    (r.type === '協力会社' && r.role && r.role !== 'only')
      ? `${r.name}${roleSuffixMap[r.role] || ''}`  // ← PR-G＋職長 の形式
      : r.name;

  const row = [
    newTimestamp, dateStr,
    companyName, staffName, siteName,
    r.type,
    nameForSave,
    toNumber(r.man), toNumber(r.overtime)
  ];
  doWithRetry(() => sheet.appendRow(row), { retries: 1, waitMs: 500 });
});

    SpreadsheetApp.flush();
    const ms = new Date() - start;
    Logger.log('%s done in %sms', FN, ms);
    return 'ok';

  } catch (e) {
    Logger.log('%s error: %s', FN, e && e.stack || e);
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// --------------------------------------------------------------------------


function toNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  const n = parseFloat(val);
  return Number.isFinite(n) ? n : 0;
}

// --------------------------------------------------------------------------
// 保存先をフォルダパスで指定できるようにする

function getOrCreateFolderByPath_(parentId, path) {
  Logger.log('--- getOrCreateFolderByPath_ 開始 ---');
  let folder = DriveApp.getFolderById(parentId);
  if (!path) return folder;
  const parts = String(path).split('/').map(s => s.trim()).filter(Boolean);
  Logger.log('フォルダパス: %s, 分割数=%s', path, parts.length);
  parts.forEach(name => {
    const it = folder.getFoldersByName(name);
    Logger.log('  フォルダ "%s" の存在確認', name);

    // 存在すればそれを取得、なければ新規作成
    folder = it.hasNext() ? it.next() : folder.createFolder(name);
  });
  Logger.log('作成先フォルダ: %s', folder.getId());
  Logger.log('--- getOrCreateFolderByPath_ 終了 ---');
  return folder;
}

// --------------------------------------------------------------------------
// Driveへ画像をアップロードする関数

function uploadImagesToDrive(meta, files) {
  Logger.log('--- uploadImagesToDrive 開始 ---');
  const SCRIPT_PROPS = PropertiesService.getScriptProperties();
  const BASE_DIR_ID  = SCRIPT_PROPS.getProperty('BASE_DIR_ID');

  Logger.log('アップロードファイル数=%s 件', files.length);
  Logger.log('メタ情報: %s', JSON.stringify(meta));

  const parentId = BASE_DIR_ID;
  if (!parentId) throw new Error('parentFolderId が指定されていません。');
  Logger.log('親フォルダID: %s', parentId);

  // === 階層を組み立てる ===
  // 例: 「美装/添付画像」 または 「揚重/添付画像」
  // meta.workType に "美装" or "揚重" が入っている想定
  const subPath = `${meta.workType}/添付画像`;

  const targetFolder = getOrCreateFolderByPath_(parentId, subPath);
  const results = [];

  files.forEach((f, i) => {
    const base64 = String(f.dataUrl).split(',')[1] || '';
    const bytes  = Utilities.base64Decode(base64);
    const mime   = f.type || MimeType.JPEG;

    // ファイル名: 出面日付_現場名_01.jpeg
    const baseName = `${meta.reportDate}_${meta.siteName}`;
    const fileName = `${baseName}_${String(i + 1).padStart(2, '0')}.jpeg`;

    // Blob 作成
    const blob = Utilities.newBlob(bytes, mime, fileName);
    const file = targetFolder.createFile(blob);

    Logger.log('アップロード完了: %s (ID=%s)',
      'https://drive.google.com/uc?id=' + file.getId(),
      file.getId()
    );
    results.push({
      id: file.getId(),
      url: 'https://drive.google.com/uc?id=' + file.getId(),
      name: file.getName()
    });
  });

  Logger.log('--- uploadImagesToDrive 終了 ---');
  return results;
}
