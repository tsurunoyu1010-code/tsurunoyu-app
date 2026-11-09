/**
 * 鶴の湯 お客様の声 — 表の読み書き（tsurunoyu-app の voices.html から呼ぶ）
 * 新しい Apps Script プロジェクトに貼り、「ウェブアプリ」としてデプロイする
 *   次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 * PIN は、チェック表アプリ（業務記録）の管理 PIN をそのまま使う
 */
const SHEET_ID = "1xGDbByDd8rAxMMos62PPOQTZ3l35htkFWXOyo7FAen8";
const TAB = "アンケート管理";
const FIRST_ROW = 4;                 // 3行目が見出し、4行目からデータ
const CHECK_API = "https://script.google.com/macros/s/AKfycbws0RsBbITD2Q6XXUiWxgpfoCdMVl8W0wxr4VgR3grVOHaBWOEpiTAmH5mwmDTD2Pe7/exec";
const WRITABLE = ["I", "J", "K", "M", "N", "R", "T"]; // 回答・副店長確認・対応状況・相良さんFB・回答確定・ロビー掲示日・備考

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = { list: list_, write: write_, ping: ping_ }[req.fn];
    if (!fn) throw new Error("unknown fn");
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function ping_(pin) { checkPin_(pin); return true; }

/* チェック表アプリに管理 PIN を問い合わせる（10分覚える） */
function checkPin_(pin) {
  pin = String(pin || "");
  if (!pin) throw new Error("PIN");
  const cache = CacheService.getScriptCache();
  const key = "pin_" + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin));
  if (cache.get(key) === "1") return;
  const res = UrlFetchApp.fetch(CHECK_API, {
    method: "post", contentType: "text/plain;charset=utf-8", muteHttpExceptions: true, followRedirects: true,
    payload: JSON.stringify({ fn: "checkPin", args: [pin] })
  });
  const j = JSON.parse(res.getContentText());
  if (!(j.ok && j.result && j.result.ok)) throw new Error("PIN");
  cache.put(key, "1", 600);
}

function sheet_() { return SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB); }

function list_(pin) {
  checkPin_(pin);
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return [];
  const vals = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 20).getDisplayValues();
  return vals.map((v, i) => ({ row: FIRST_ROW + i, v: v })).filter(x => x.v[0]);
}

/* updates = { I: "…", J: "確認済み", … }。行の管理IDが id と同じときだけ書く */
function write_(pin, row, id, updates) {
  checkPin_(pin);
  row = Number(row);
  if (!(row >= FIRST_ROW)) throw new Error("row");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    if (sh.getRange(row, 1).getDisplayValue() !== String(id)) throw new Error("moved");
    Object.keys(updates || {}).forEach(col => {
      if (WRITABLE.indexOf(col) < 0) throw new Error("col " + col);
      const val = String(updates[col] == null ? "" : updates[col]).slice(0, 2000);
      sh.getRange(col + row).setValue(val);
    });
    SpreadsheetApp.flush();
    return sh.getRange(row, 1, 1, 20).getDisplayValues()[0];
  } finally {
    lock.releaseLock();
  }
}
