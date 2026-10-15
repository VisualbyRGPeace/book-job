// Dán vào Google Sheet: Extensions > Apps Script
const SHEET_NAME = "Sheet1";      // đổi đúng tên tab chứa bảng job
const FIRST_DATA_ROW = 3;         // dòng dữ liệu đầu tiên (như ảnh: dòng 3)
const DEFAULT_STATUS = "Đã nhận"; // trạng thái mặc định khi book job
const TZ = "Asia/Ho_Chi_Minh";

// Cột: A Task | B Title | C Deadline | D Completed | E Status | F Designer | G Quantity | H Assignment | I,J Score (giữ công thức) | K Date&Time
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

    // tìm dòng trống đầu tiên theo cột B (không ghi đè công thức / dropdown)
    const colB = sh.getRange(FIRST_DATA_ROW, 2, Math.max(sh.getLastRow() - FIRST_DATA_ROW + 1, 1), 1).getValues();
    let n = colB.findIndex(r => r[0] === "");
    const row = FIRST_DATA_ROW + (n === -1 ? colB.length : n);

    const [y, m, day] = d.deadline.split("-").map(Number);
    sh.getRange(row, 1, 1, 8).setValues([[
      d.task, d.title, new Date(y, m - 1, day), false,
      DEFAULT_STATUS, d.designer || "", d.qty, d.assign
    ]]);
    sh.getRange(row, 11).setValue(new Date()).setNumberFormat("dd/MM/yyyy H:mm:ss");

    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const limit = Number((e.parameter || {}).limit) || 15;
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const last = sh.getLastRow();
  if (last < FIRST_DATA_ROW) return out({ rows: [] });
  const vals = sh.getRange(FIRST_DATA_ROW, 1, last - FIRST_DATA_ROW + 1, 8).getValues();
  const rows = vals.filter(r => r[1] !== "").reverse().slice(0, limit).map(r => ({
    task: r[0], title: r[1],
    deadline: r[2] instanceof Date ? Utilities.formatDate(r[2], TZ, "dd/MM/yyyy") : r[2],
    status: r[4], designer: r[5], qty: r[6]
  }));
  return out({ rows });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
