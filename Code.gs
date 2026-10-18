/**
 * GOOGLE APPS SCRIPT — lưu đăng ký vào Google Sheets.
 *
 * 1) Điền ID bảng tính vào SPREADSHEET_ID.
 * 2) Tên trang tính phải khớp SHEET_NAME.
 * 3) Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: Anyone (nếu tài khoản cho phép).
 * 4) Dùng URL kết thúc bằng /exec trong action của index.html.
 */

const SPREADSHEET_ID = "1z5Cbh5fvQc31ODj5zTyu0YfPSQ4QmLs-8QVIpJ86LeM";
const SHEET_NAME = "DangKy";

function doGet() {
  return HtmlService.createHtmlOutput(
    "<!doctype html><html><body>Dịch vụ đăng ký đang hoạt động.</body></html>"
  );
}

function doPost(e) {
  let result = {
    type: "registration-result",
    ok: false,
    message: "Không thể lưu đăng ký. Vui lòng thử lại."
  };

  try {
    const p = (e && e.parameter) ? e.parameter : {};

    // Honeypot: bot điền trường ẩn thì bỏ qua.
    if (String(p.website || "").trim() !== "") {
      result.ok = true;
      result.message = "";
      return responseHtml_(result);
    }

    const hoTen = clean_(p.hoTen);
    const dienThoai = clean_(p.dienThoai);

    if (!hoTen || !dienThoai) {
      throw new Error("Vui lòng nhập họ tên và số điện thoại.");
    }

    if (hoTen.length > 120 || dienThoai.length > 25) {
      throw new Error("Thông tin vượt quá độ dài cho phép.");
    }

    if (!SPREADSHEET_ID || SPREADSHEET_ID.includes("/") || SPREADSHEET_ID.includes("?")) {
      throw new Error("SPREADSHEET_ID phải là ID bảng tính, không phải URL.");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME);

      if (!sheet) {
        throw new Error('Không tìm thấy trang tính "' + SHEET_NAME + '".');
      }

      // Định dạng số điện thoại là văn bản để giữ số 0 ở đầu.
      const row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, 3).setValues([[
        new Date(),
        safeCell_(hoTen),
        safeCell_(dienThoai)
      ]]);
      sheet.getRange(row, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    } finally {
      lock.releaseLock();
    }

    result.ok = true;
    result.message = "";
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    // Không gửi chi tiết nội bộ ra trình duyệt.
    result.ok = false;
    result.message = "Chưa lưu được đăng ký. Vui lòng kiểm tra cấu hình hoặc liên hệ nhà trường.";
  }

  return responseHtml_(result);
}

function clean_(value) {
  return String(value == null ? "" : value).trim();
}

/**
 * Giảm nguy cơ công thức bảng tính khi dữ liệu bắt đầu bằng ký tự nguy hiểm.
 * Dữ liệu vẫn được lưu như văn bản.
 */
function safeCell_(value) {
  if (/^[=+\-@]/.test(value)) {
    return "'" + value;
  }
  return value;
}

function responseHtml_(result) {
  const json = JSON.stringify(result);
  const safeJson = json
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026");

  const html = `<!doctype html>
<html>
<head><meta charset="utf-8"></head>
<body>
<script>
  window.parent.postMessage(${JSON.stringify(safeJson)}, "*");
</script>
Đã xử lý yêu cầu.
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Chạy thủ công một lần trong Apps Script để kiểm tra:
 * - ID bảng tính đúng
 * - có trang tính DangKy
 * Hàm này không thêm dòng dữ liệu.
 */
function kiemTraCauHinh() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.includes("/") || SPREADSHEET_ID.includes("?")) {
    throw new Error("SPREADSHEET_ID phải là ID bảng tính, không phải URL.");
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Không tìm thấy trang tính "' + SHEET_NAME + '".');
  }
  Logger.log("OK: " + ss.getName() + " / " + sheet.getName());
}
