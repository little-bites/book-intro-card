const DEADLINE = new Date("2026-11-19T23:59:00+09:00");
const HEADERS = ["送信日時","ニックネーム", ...Array.from({length:10},(_,i)=>"Q"+(i+1))];

function doPost(e) {
  try {
    if (new Date() > DEADLINE) return output({ok:false, message:"受付期間は終了しました。"});
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw) return output({ok:false, message:"データがありません。"});
    const data = JSON.parse(raw);
    const nickname = String(data.nickname || "").trim().slice(0,30);
    if (!nickname) return output({ok:false, message:"ニックネームを入力してください。"});
    if (!Array.isArray(data.answers) || data.answers.length !== 10) return output({ok:false, message:"回答形式が正しくありません。"});
    const sheet = getResponseSheet();
    const row = [new Date(), nickname, ...data.answers.map(a => {
      if (typeof a.answer === "string") return a.answer.slice(0,2000);
      return Object.entries(a.answer || {}).map(([label, value]) => label + "：" + String(value || "")).join(" / ").slice(0,2000);
    })];
    sheet.appendRow(row);
    return output({ok:true});
  } catch (err) {
    return output({ok:false, message:String(err)});
  }
}
function doGet() { return output({ok:true, message:"BOOK QUEST受付API"}); }
function getResponseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("回答一覧");
  if (!sheet) sheet = ss.insertSheet("回答一覧");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function output(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
